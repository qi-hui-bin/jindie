from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, dataclass, replace
from typing import Mapping

from .accounting import (
    AccountingError, Assignment, Auxiliary, Entry, Ledger, Source, Voucher, edit_entries, manual_voucher, merge_vouchers, review,
    propose, split_voucher, validate_voucher,
)
from .kingdee_export import ExportOptions, export_workbook, reconciliation_workbook
from .models import Issue, Period, StatementTransaction, canonical_json, fingerprint
from .storage import StaleBookSetup, StatementStore, StorageError
from .voucher_preview import parse_config, plan_batch, transaction_from_json


class StaleReview(StorageError):
    """The operation was submitted against a previous revision."""


SKIP_REASONS = {"posted": "历史已入账", "other_account": "已在同账套另一账户入账", "non_transaction": "确认为非交易行"}
PENDING_GROUPS = {
    "unmatched": ("未匹配规则", ("unmatched",)),
    "conflict": ("规则冲突", ("rule_conflict",)),
    "coding": ("科目或辅助编码", ("unknown_subject", "missing_auxiliary", "invalid_auxiliary", "unknown_auxiliary")),
    "voucher": ("其他制证问题", ("empty_summary", "quantity_required")),
}
MANUAL_CODES = frozenset(code for _, codes in PENDING_GROUPS.values() for code in codes)


def can_resolve_manually(exclusion: dict) -> bool:
    codes = {issue["code"] for issue in exclusion["issues"]}
    return not exclusion.get("skipped") and bool(codes) and codes <= MANUAL_CODES


def _pack(voucher: Voucher) -> str:
    return canonical_json({
        "sources": [
            {"transaction_id": source.transaction_id, "book_id": source.book_id, "account_id": source.account_id,
             "period": str(source.period), "batch_id": source.batch_id, "sheet": source.sheet,
             "row_number": source.row_number, "transaction": source.transaction.payload()}
            for source in voucher.sources
        ],
        "entries": [asdict(entry) for entry in voucher.entries],
        "rule_versions": voucher.rule_versions,
    })


def _unpack(payload: str, digest: str | None) -> Voucher:
    data = json.loads(payload)
    sources = tuple(Source(**{
        **source, "period": Period.from_text(source["period"]),
        "transaction": transaction_from_json(source["transaction"]),
    }) for source in data["sources"])
    entries = tuple(Entry(
        Assignment(item["assignment"]["subject"], tuple(Auxiliary(**aux) for aux in item["assignment"]["auxiliary"])),
        item["summary"], item["side"], item["amount_minor"], tuple(item["transaction_ids"]),
    ) for item in data["entries"])
    return Voucher(sources, entries, tuple((key, version) for key, version in data["rule_versions"]), digest)


@dataclass(frozen=True)
class Draft:
    id: int
    token: str
    voucher: Voucher
    export_id: int | None
    original_export_id: int | None
    original_number: int | None


@dataclass(frozen=True)
class PendingAccountingRow:
    book_id: str
    account_id: str
    batch_id: int
    period: str
    sheet: str
    row_number: int
    transaction: StatementTransaction
    issues: tuple[Issue, ...]
    can_make_draft: bool
    setup_version: int | None


@dataclass(frozen=True)
class PendingRecalculation:
    source: Source
    origin_version: int
    revision: int
    ledger: Ledger
    voucher: Voucher | None
    issues: tuple[Issue, ...]
    prior_issues: tuple[Issue, ...]


def _draft(row: sqlite3.Row) -> Draft:
    return Draft(row["id"], fingerprint([row["payload_json"], row["review_digest"]]),
                 _unpack(row["payload_json"], row["review_digest"]), row["export_id"], row["original_export_id"],
                 row["original_number"])


def _draft_rows(conn: sqlite3.Connection, batch_id: int, draft_id: int | None = None) -> list[sqlite3.Row]:
    return conn.execute(
        """SELECT d.*, e.export_id, r.original_export_id, r.original_number FROM voucher_drafts d
           LEFT JOIN exported_vouchers e ON e.draft_id = d.id
           LEFT JOIN voucher_retries r ON r.draft_id = d.id
           WHERE d.batch_id = ? AND d.superseded = 0 AND (? IS NULL OR d.id = ?) ORDER BY d.id""",
        (batch_id, draft_id, draft_id),
    ).fetchall()


def _frozen_ledger(record: sqlite3.Row) -> Ledger:
    return parse_config(json.loads(record["book_config_json"]))[0]


def _skips(batch: sqlite3.Row) -> dict[int, dict]:
    return {int(row): details for row, details in json.loads(batch["skips_json"]).items()}


def review_exclusions(batch: sqlite3.Row) -> list[dict]:
    skips = _skips(batch)
    remaining = [item for item in json.loads(batch["exclusions_json"]) if item["row_number"] not in skips]
    remaining.extend({"row_number": row, "skipped": True,
                      "issues": [{"code": "confirmed_skip", "message": f"{SKIP_REASONS[item['reason']]}：{item['note']}"}]}
                     for row, item in skips.items())
    return sorted(remaining, key=lambda item: item["row_number"])


def _log(conn: sqlite3.Connection, batch_id: int, kind: str, details: dict) -> None:
    conn.execute("INSERT INTO review_actions(batch_id, kind, details_json) VALUES (?, ?, ?)",
                 (batch_id, kind, canonical_json(details)))
    conn.execute("UPDATE review_batches SET revision = revision + 1 WHERE batch_id = ?", (batch_id,))


def _new_draft(conn: sqlite3.Connection, batch_id: int, voucher: Voucher) -> int:
    return conn.execute("INSERT INTO voucher_drafts(batch_id, payload_json) VALUES (?, ?) RETURNING id",
                        (batch_id, _pack(voucher))).fetchone()["id"]


def _manual_source(conn: sqlite3.Connection, batch: sqlite3.Row, row_number: int) -> tuple[Source, dict]:
    exclusion = next((item for item in review_exclusions(batch) if item["row_number"] == row_number), None)
    if exclusion is None or not can_resolve_manually(exclusion):
        raise StorageError("该行不属于可手工制证的规则异常，或已处理、已跳过。")
    row = conn.execute(
        """SELECT r.transaction_id, r.parsed_json FROM source_rows r
           LEFT JOIN exported_sources x ON x.transaction_id = r.transaction_id
           WHERE r.batch_id = ? AND r.row_number = ? AND r.match_status = 'new'
             AND r.issues_json = '[]' AND r.parsed_json IS NOT NULL
             AND r.transaction_id IS NOT NULL AND x.transaction_id IS NULL""",
        (batch["id"], row_number),
    ).fetchone()
    if row is None:
        raise StorageError("流水未完成解析、存在查重问题或已经导出，不能手工制证。")
    transaction = transaction_from_json(json.loads(row["parsed_json"]))
    if transaction.currency != "CNY":
        raise AccountingError("非人民币流水暂不支持手工制证。")
    return Source(row["transaction_id"], batch["book_id"], batch["account_id"], Period.from_text(batch["period"]),
                  batch["id"], batch["sheet"], row_number, transaction), exclusion


class VoucherWorkflow(StatementStore):
    def pending_accounting(
        self, book_ids: tuple[str, ...], *, period: str | None, group: str, limit: int, offset: int,
    ) -> list[PendingAccountingRow]:
        if not book_ids:
            return []
        codes = PENDING_GROUPS[group][1] if group else tuple(sorted(MANUAL_CODES))
        book_slots = ", ".join("?" for _ in book_ids)
        code_slots = ", ".join("?" for _ in codes)
        with self._connection() as conn:
            rows = conn.execute(
                f"""SELECT b.book_id, b.account_id, b.id AS batch_id, b.period, b.sheet, b.setup_version,
                           r.row_number, r.parsed_json, exception.value AS exclusion
                    FROM import_batches b JOIN review_batches q ON q.batch_id = b.id
                    JOIN json_each(q.exclusions_json) exception
                    JOIN source_rows r ON r.batch_id = b.id
                      AND r.row_number = json_extract(exception.value, '$.row_number')
                    LEFT JOIN exported_sources x ON x.transaction_id = r.transaction_id
                    WHERE b.book_id IN ({book_slots}) AND (? IS NULL OR b.period = ?)
                      AND r.match_status = 'new' AND r.issues_json = '[]' AND r.parsed_json IS NOT NULL
                      AND r.transaction_id IS NOT NULL AND x.transaction_id IS NULL
                      AND NOT EXISTS (SELECT 1 FROM json_each(q.skips_json) skip
                                      WHERE skip.key = CAST(r.row_number AS TEXT))
                      AND EXISTS (SELECT 1 FROM json_each(exception.value, '$.issues') issue
                                  WHERE json_extract(issue.value, '$.code') IN ({code_slots}))
                    ORDER BY b.id DESC, r.row_number LIMIT ? OFFSET ?""",
                (*book_ids, period, period, *codes, limit, offset),
            ).fetchall()
        result = []
        for row in rows:
            exclusion = json.loads(row["exclusion"])
            transaction = transaction_from_json(json.loads(row["parsed_json"]))
            result.append(PendingAccountingRow(
                row["book_id"], row["account_id"], row["batch_id"], row["period"], row["sheet"], row["row_number"],
                transaction, tuple(Issue(**data) for data in exclusion["issues"]),
                transaction.currency == "CNY" and can_resolve_manually(exclusion), row["setup_version"],
            ))
        return result

    def _batch(self, conn: sqlite3.Connection, batch_id: int, book_id: str) -> sqlite3.Row:
        row = conn.execute(
            """SELECT b.*, r.revision, r.book_config_json, r.exclusions_json, r.skips_json,
                      link.origin_batch_id, link.origin_row_number
               FROM import_batches b JOIN review_batches r ON r.batch_id = b.id
               LEFT JOIN recalculated_rows link ON link.target_batch_id = b.id
               WHERE b.id = ? AND b.book_id = ?""", (batch_id, book_id),
        ).fetchone()
        if row is None:
            raise StorageError("未找到该账套的复核批次。")
        return row

    def prepare(self, batch_id: int, config: dict) -> None:
        ledger, rules = parse_config(config)
        self.initialize()
        with self._connection() as conn:
            batch = conn.execute(
                """SELECT b.setup_version, r.batch_id AS prepared, v.config_json
                   FROM import_batches b LEFT JOIN review_batches r ON r.batch_id = b.id
                   LEFT JOIN book_setup_versions v ON v.book_id = b.book_id AND v.version = b.setup_version
                   WHERE b.id = ? AND b.book_id = ?""", (batch_id, ledger.book_id),
            ).fetchone()
            if batch is None:
                raise StorageError("未找到该账套的流水批次。")
            if batch["prepared"] is not None:
                return
            if batch["setup_version"] is not None:
                if batch["config_json"] is None:
                    raise StorageError("找不到批次最初使用的账套配置版本，不能生成凭证。")
                pinned = json.loads(batch["config_json"])
                config = {"ledger": pinned["ledger"], "rules": pinned["rules"]}
                ledger, rules = parse_config(config)
        _, proposal = plan_batch(self, batch_id, ledger, rules)
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            created = conn.execute(
                """INSERT INTO review_batches(batch_id, book_config_json, exclusions_json) VALUES (?, ?, ?)
                   ON CONFLICT DO NOTHING""",
                (batch_id, canonical_json(config), canonical_json([asdict(item) for item in proposal.exclusions])),
            )
            if created.rowcount:
                conn.executemany("INSERT INTO voucher_drafts(batch_id, payload_json) VALUES (?, ?)",
                                 ((batch_id, _pack(voucher)) for voucher in proposal.vouchers))

    def review_data(self, batch_id: int, book_id: str) -> tuple[sqlite3.Row, Ledger, list[Draft], list[sqlite3.Row]]:
        with self._connection() as conn:
            conn.execute("BEGIN")
            batch = self._batch(conn, batch_id, book_id)
            drafts = [_draft(row) for row in _draft_rows(conn, batch_id)]
            exports = conn.execute(
                """SELECT id, created_at, voucher_word FROM voucher_exports
                   WHERE batch_id = ? ORDER BY id DESC""", (batch_id,),
            ).fetchall()
            return batch, _frozen_ledger(batch), drafts, exports

    def change_draft(
        self, book_id: str, batch_id: int, draft_id: int, token: str, *,
        edits: Mapping[int, tuple[str, Assignment]] | None = None,
    ) -> None:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            ledger = _frozen_ledger(batch)
            rows = _draft_rows(conn, batch_id, draft_id)
            if not rows or rows[0]["export_id"] is not None:
                raise StorageError("该凭证不存在或已导出锁定。")
            draft = _draft(rows[0])
            if draft.token != token:
                raise StaleReview("凭证已经变更，请刷新并重新核对。")
            voucher = draft.voucher
            skips = _skips(batch)
            if any(source.row_number in skips for source in voucher.sources):
                raise StorageError("该凭证含已跳过流水，请先撤回跳过。")
            if edits is None:
                voucher = review(voucher, ledger)
            else:
                if not edits or not edits.keys() <= set(range(len(voucher.entries))):
                    raise AccountingError("没有可保存的分录修改。")
                entries = tuple(replace(entry, summary=edits[i][0], assignment=edits[i][1]) if i in edits else entry
                                for i, entry in enumerate(voucher.entries))
                voucher, _ = edit_entries(voucher, entries, ledger)
            conn.execute("UPDATE voucher_drafts SET payload_json = ?, review_digest = ? WHERE id = ?",
                         (_pack(voucher), voucher.review_digest, draft_id))
            conn.execute("UPDATE review_batches SET revision = revision + 1 WHERE batch_id = ?", (batch_id,))

    def manual_details(self, book_id: str, batch_id: int, row_number: int) -> tuple[int, Ledger, Source, dict]:
        with self._connection() as conn:
            conn.execute("BEGIN")
            batch = self._batch(conn, batch_id, book_id)
            source, exclusion = _manual_source(conn, batch, row_number)
            return batch["revision"], _frozen_ledger(batch), source, exclusion

    def create_manual(
        self, book_id: str, batch_id: int, revision: int, row_number: int,
        assignment: Assignment, summary: str, reason: str,
    ) -> int:
        if not reason.strip() or len(reason) > 500:
            raise AccountingError("请填写手工制证依据（最多 500 字）。")
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            if batch["revision"] != revision:
                raise StaleReview("批次已经变更，请刷新后操作。")
            source, original = _manual_source(conn, batch, row_number)
            draft_id = _new_draft(conn, batch_id, manual_voucher(source, _frozen_ledger(batch), assignment, summary))
            remaining = [item for item in json.loads(batch["exclusions_json"]) if item["row_number"] != row_number]
            conn.execute("UPDATE review_batches SET exclusions_json = ? WHERE batch_id = ?",
                         (canonical_json(remaining), batch_id))
            _log(conn, batch_id, "manual_voucher", {"row": row_number, "draft": draft_id, "reason": reason.strip(),
                                                    "prior_issues": original["issues"]})
            return draft_id

    def _recalculation(
        self, conn: sqlite3.Connection, batch: sqlite3.Row, row_number: int, *, expected_version: int | None = None,
    ) -> tuple[PendingRecalculation, dict]:
        latest = conn.execute(
            "SELECT version, config_json FROM book_setup_versions WHERE book_id = ? ORDER BY version DESC LIMIT 1",
            (batch["book_id"],),
        ).fetchone()
        if expected_version is not None and (latest is None or latest["version"] != expected_version):
            raise StaleBookSetup("账套配置已经更新，请重新试算。")
        if latest is None or batch["setup_version"] is None or latest["version"] <= batch["setup_version"]:
            raise AccountingError("当前没有比原批次更新的已发布配置。")
        source, exclusion = _manual_source(conn, batch, row_number)
        config = json.loads(latest["config_json"])
        snapshot = {"ledger": config["ledger"], "rules": config["rules"]}
        if not any(account["id"] == source.account_id for account in config["web"]["accounts"]):
            raise AccountingError("该银行账户在最新版配置中已停用，不能重算。")
        ledger, rules = parse_config(snapshot)
        proposal = propose((source,), ledger, rules)
        return PendingRecalculation(
            source, batch["setup_version"], batch["revision"], ledger,
            next(iter(proposal.vouchers), None),
            tuple(issue for exclusion in proposal.exclusions for issue in exclusion.issues),
            tuple(Issue(**item) for item in exclusion["issues"]),
        ), snapshot

    def preview_recalculation(self, book_id: str, batch_id: int, row_number: int) -> PendingRecalculation:
        with self._connection() as conn:
            conn.execute("BEGIN")
            batch = self._batch(conn, batch_id, book_id)
            return self._recalculation(conn, batch, row_number)[0]

    def apply_recalculation(
        self, book_id: str, batch_id: int, row_number: int, *, revision: int, version: int,
    ) -> tuple[int, int]:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            if batch["revision"] != revision:
                raise StaleReview("原批次已变更，请回到待确认清单确认该行仍未处理，再重新试算。")
            trial, snapshot = self._recalculation(conn, batch, row_number, expected_version=version)
            if trial.voucher is None:
                raise AccountingError("新版规则未能生成有效凭证，不能转入新批次。")
            target_id = conn.execute(
                """INSERT INTO import_batches (
                       book_id, account_id, period, profile_id, profile_version, file_hash, filename,
                       sheet, header_row, options_json, warnings_json, setup_version
                   ) SELECT book_id, account_id, period, profile_id, profile_version, file_hash, filename,
                            sheet, header_row, ?, warnings_json, ? FROM import_batches WHERE id = ?
                   RETURNING id""",
                (canonical_json({"recalculated_from": [batch_id, row_number]}), version, batch_id),
            ).fetchone()["id"]
            conn.execute(
                """INSERT INTO source_rows
                       (batch_id, row_number, content_hash, raw_json, parsed_json, issues_json, match_status, transaction_id)
                   SELECT ?, row_number, content_hash, raw_json, parsed_json, issues_json, match_status, transaction_id
                   FROM source_rows WHERE batch_id = ? AND row_number = ?""", (target_id, batch_id, row_number),
            )
            conn.execute("INSERT INTO review_batches(batch_id, book_config_json, exclusions_json) VALUES (?, ?, '[]')",
                         (target_id, canonical_json(snapshot)))
            conn.execute("INSERT INTO recalculated_rows VALUES (?, ?, ?)", (batch_id, row_number, target_id))
            voucher = replace(trial.voucher, sources=(replace(trial.source, batch_id=target_id),))
            draft_id = _new_draft(conn, target_id, voucher)
            exclusions = json.loads(batch["exclusions_json"])
            for index, item in enumerate(exclusions):
                if item["row_number"] == row_number:
                    exclusions[index] = {
                        "row_number": row_number, "recalculated_batch_id": target_id,
                        "issues": [{"code": "recalculated", "message": f"已转至重算批次 #{target_id}，原批次不再制证。"}],
                    }
                    break
            conn.execute("UPDATE review_batches SET exclusions_json = ? WHERE batch_id = ?",
                         (canonical_json(exclusions), batch_id))
            _log(conn, batch_id, "recalculate", {
                "row": row_number, "target_batch": target_id, "draft": draft_id, "setup_version": version,
                "prior_issues": [asdict(issue) for issue in trial.prior_issues],
            })
            return target_id, draft_id

    def set_skip(
        self, book_id: str, batch_id: int, revision: int, row_number: int, reason: str | None, note: str = "",
    ) -> None:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            if batch["revision"] != revision:
                raise StaleReview("批次已经变更，请刷新后操作。")
            source = conn.execute("SELECT 1 FROM source_rows WHERE batch_id = ? AND row_number = ?",
                                  (batch_id, row_number)).fetchone()
            if source is None:
                raise StorageError("该来源行不属于当前批次。")
            if conn.execute("SELECT 1 FROM recalculated_rows WHERE origin_batch_id = ? AND origin_row_number = ?",
                            (batch_id, row_number)).fetchone():
                raise StorageError("该来源行已转至重算批次，请在重算批次中处理。")
            skips = _skips(batch)
            if reason is None:
                if row_number not in skips:
                    raise StorageError("该来源行未被跳过。")
                detail = skips.pop(row_number)
            else:
                if row_number in skips or reason not in SKIP_REASONS or not note.strip() or len(note) > 500:
                    raise StorageError("请选择跳过原因，并填写核对依据（最多 500 字）。")
                for row in _draft_rows(conn, batch_id):
                    draft = _draft(row)
                    if not any(s.row_number == row_number for s in draft.voucher.sources):
                        continue
                    if draft.export_id is not None or draft.original_export_id is not None:
                        raise StorageError("已导出和重导修订中的来源流水不能跳过。")
                    if len(draft.voucher.sources) > 1:
                        raise StorageError("该行属于多笔合并凭证，请先拆分再跳过。")
                    conn.execute("UPDATE voucher_drafts SET review_digest = NULL WHERE id = ?", (draft.id,))
                detail = skips[row_number] = {"reason": reason, "note": note.strip()}
            conn.execute("UPDATE review_batches SET skips_json = ? WHERE batch_id = ?", (canonical_json(skips), batch_id))
            _log(conn, batch_id, "skip" if reason is not None else "undo_skip", {"row": row_number, **detail})

    def regroup(self, book_id: str, batch_id: int, revision: int, draft_ids: tuple[int, ...], *, split: bool) -> list[int]:
        if len(set(draft_ids)) != len(draft_ids) or (len(draft_ids) != 1 if split else len(draft_ids) < 2):
            raise AccountingError("请选择一张待拆分凭证，或至少两张不同凭证进行合并。")
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            if batch["revision"] != revision:
                raise StaleReview("批次已经变更，请刷新后操作。")
            found = {row["id"]: _draft(row) for row in _draft_rows(conn, batch_id) if row["id"] in draft_ids}
            skips = _skips(batch)
            if len(found) != len(draft_ids) or any(
                item.export_id is not None or item.original_export_id is not None or
                any(s.row_number in skips for s in item.voucher.sources) for item in found.values()
            ):
                raise StorageError("只能拆并本批未导出、未跳过且不属于重导修订的凭证。")
            ledger = _frozen_ledger(batch)
            old = tuple(found[id].voucher for id in draft_ids)
            new = split_voucher(old[0], ledger) if split else (merge_vouchers(old, ledger),)
            conn.executemany("UPDATE voucher_drafts SET superseded = 1, review_digest = NULL WHERE id = ?",
                             ((id,) for id in draft_ids))
            created = [_new_draft(conn, batch_id, voucher) for voucher in new]
            _log(conn, batch_id, "split" if split else "merge", {"old": draft_ids, "new": created})
            return created

    def start_retry(self, book_id: str, export_id: int, number: int) -> tuple[int, int]:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute(
                """SELECT x.batch_id, e.snapshot_json, e.import_status, r.draft_id FROM exported_vouchers e
                   JOIN voucher_exports x ON x.id = e.export_id
                   LEFT JOIN voucher_retries r ON r.original_export_id = e.export_id AND r.original_number = e.number
                   WHERE e.export_id = ? AND e.number = ? AND e.book_id = ?""", (export_id, number, book_id),
            ).fetchone()
            if row is None or row["import_status"] != "not_imported":
                raise StorageError("只有选择「已核实未导入」的凭证才能发起修订重导。")
            batch_id = row["batch_id"]
            if row["draft_id"] is not None:
                return batch_id, row["draft_id"]
            draft_id = _new_draft(conn, batch_id, _unpack(row["snapshot_json"], None))
            conn.execute("INSERT INTO voucher_retries(draft_id, original_export_id, original_number) VALUES (?, ?, ?)",
                         (draft_id, export_id, number))
            _log(conn, batch_id, "retry", {"export": export_id, "number": number, "draft": draft_id})
            return batch_id, draft_id

    def create_export(
        self, book_id: str, batch_id: int, revision: int, request_id: str, template: bytes, options: ExportOptions,
        *, allow_partial: bool = False, retry_draft_id: int | None = None,
    ) -> int:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            batch = self._batch(conn, batch_id, book_id)
            option_values = {
                "revision": revision, "partial": allow_partial, "first_number": options.first_number,
                "used_numbers": sorted(options.used_numbers), "text_number": options.text_number,
                "currency_code": options.currency_code, "period": str(options.period), "word": options.voucher_word,
            }
            if retry_draft_id is not None:
                option_values["retry_draft"] = retry_draft_id
            request_options = canonical_json(option_values)
            existing = conn.execute(
                "SELECT id, options_json FROM voucher_exports WHERE batch_id = ? AND request_id = ?", (batch_id, request_id),
            ).fetchone()
            if existing is not None:
                if existing["options_json"] != request_options:
                    raise StaleReview("该导出请求已使用。请刷新页面后重新选择范围和编号。")
                return existing["id"]
            if batch["revision"] != revision:
                raise StaleReview("批次中有凭证发生变化，请刷新后重新确认导出范围。")
            if str(options.period) != batch["period"]:
                raise StorageError("导出期间与批次不一致。")
            ledger = _frozen_ledger(batch)
            exclusions = review_exclusions(batch) if retry_draft_id is None else []
            skips = _skips(batch)
            selected: list[Draft] = []
            for row in _draft_rows(conn, batch_id, retry_draft_id):
                draft = _draft(row)
                if draft.export_id is not None:
                    continue
                if (draft.original_export_id is None) != (retry_draft_id is None):
                    continue
                voucher = draft.voucher
                if any(source.row_number in skips for source in voucher.sources):
                    continue
                if voucher.review_digest == voucher.digest(ledger) and not validate_voucher(voucher, ledger):
                    selected.append(draft)
                else:
                    exclusions.extend({"row_number": source.row_number,
                                       "issues": [{"code": "unreviewed", "message": "关联凭证未复核或未通过校验。"}]}
                                      for source in voucher.sources)
            if not selected:
                raise StorageError("没有通过校验且尚未导出的已复核凭证。")
            if exclusions and not allow_partial:
                raise StorageError("本批有排除项；请核对明细后明确确认只导出已复核凭证。")
            scope = book_id, batch["period"], options.voucher_word
            local_numbers = frozenset(row["number"] for row in conn.execute(
                "SELECT number FROM exported_vouchers WHERE book_id = ? AND period = ? AND voucher_word = ?", scope,
            ))
            self._check_sources(conn, batch_id, selected)
            result = export_workbook(template, (draft.voucher for draft in selected), ledger,
                                     replace(options, used_numbers=options.used_numbers | local_numbers), preview=False)
            source_rows = conn.execute(
                "SELECT row_number, parsed_json FROM source_rows WHERE batch_id = ? ORDER BY row_number", (batch_id,),
            )
            reconciliation = reconciliation_workbook(
                result, ledger, options,
                ((row["row_number"], json.loads(row["parsed_json"] or "null") or {}) for row in source_rows),
                exclusions,
            )
            export_id = conn.execute(
                """INSERT INTO voucher_exports(batch_id, request_id, book_id, period, voucher_word, template_hash,
                                               options_json, exclusions_json, workbook, reconciliation)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id""",
                (batch_id, request_id, *scope, options.verified_template_hash, request_options,
                 canonical_json(exclusions), result.data, reconciliation),
            ).fetchone()["id"]
            draft_by_source = {draft.voucher.sources[0].transaction_id: draft for draft in selected}
            for item in result.numbered:
                draft = draft_by_source[item.voucher.sources[0].transaction_id]
                conn.execute(
                    """INSERT INTO exported_vouchers(export_id, draft_id, book_id, period, voucher_word, number, snapshot_json)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (export_id, draft.id, *scope, item.number, _pack(item.voucher)),
                )
                for source in item.voucher.sources:
                    if draft.original_export_id is None:
                        conn.execute("INSERT INTO exported_sources(transaction_id, export_id, number) VALUES (?, ?, ?)",
                                     (source.transaction_id, export_id, item.number))
                    else:
                        conn.execute("UPDATE exported_sources SET export_id = ?, number = ? WHERE transaction_id = ?",
                                     (export_id, item.number, source.transaction_id))
            conn.execute("UPDATE review_batches SET revision = revision + 1 WHERE batch_id = ?", (batch_id,))
            return export_id

    def _check_sources(self, conn: sqlite3.Connection, batch_id: int, drafts: list[Draft]) -> None:
        valid = {row["transaction_id"]: (row["export_id"], row["number"], row["import_status"]) for row in conn.execute(
            """SELECT r.transaction_id, x.export_id, x.number, e.import_status
               FROM source_rows r LEFT JOIN exported_sources x ON x.transaction_id = r.transaction_id
               LEFT JOIN exported_vouchers e ON e.export_id = x.export_id AND e.number = x.number
               LEFT JOIN recalculated_rows link ON link.origin_batch_id = r.batch_id
                 AND link.origin_row_number = r.row_number
               WHERE r.batch_id = ? AND r.match_status = 'new' AND r.issues_json = '[]' AND link.target_batch_id IS NULL""",
            (batch_id,),
        )}
        for draft in drafts:
            expected = (draft.original_export_id, draft.original_number, "not_imported" if draft.original_export_id else None)
            if any(valid.get(source.transaction_id) != expected for source in draft.voucher.sources):
                raise StorageError("部分来源流水不属于本批、仍有异常或已导出且未获重导确认。")

    def export_file(self, book_id: str, export_id: int, *, reconciliation: bool = False) -> bytes:
        with self._connection() as conn:
            row = conn.execute("SELECT workbook, reconciliation FROM voucher_exports WHERE id = ? AND book_id = ?",
                               (export_id, book_id)).fetchone()
            if row is None:
                raise StorageError("未找到该账套的导出文件。")
            content = row["reconciliation" if reconciliation else "workbook"]
            if content is None:
                raise StorageError("此历史导出没有保存核对底稿。")
            return content

    def export_data(self, book_id: str, export_id: int) -> tuple[sqlite3.Row, list[tuple[sqlite3.Row, Voucher]]]:
        with self._connection() as conn:
            row = conn.execute(
                """SELECT id, batch_id, period, voucher_word, created_at, exclusions_json,
                          reconciliation IS NOT NULL AS has_reconciliation FROM voucher_exports
                   WHERE id = ? AND book_id = ?""", (export_id, book_id),
            ).fetchone()
            if row is None:
                raise StorageError("未找到该账套的导出记录。")
            entries = conn.execute(
                """SELECT e.number, e.snapshot_json, e.import_status, e.import_note, r.draft_id AS retry_draft_id,
                          replacement.export_id AS replacement_export_id, previous.original_export_id, previous.original_number
                   FROM exported_vouchers e
                   LEFT JOIN voucher_retries r ON r.original_export_id = e.export_id AND r.original_number = e.number
                   LEFT JOIN exported_vouchers replacement ON replacement.draft_id = r.draft_id
                   LEFT JOIN voucher_retries previous ON previous.draft_id = e.draft_id
                   WHERE e.export_id = ? ORDER BY e.number""", (export_id,),
            ).fetchall()
            return row, [(entry, _unpack(entry["snapshot_json"], None)) for entry in entries]

    def record_import(self, book_id: str, export_id: int, number: int, status: str, note: str) -> None:
        if status not in ("imported", "failed", "not_imported") or len(note) > 500 or (status != "imported" and not note.strip()):
            raise StorageError("请选择有效的导入状态；失败或未入账须填写原因（最多 500 字）。")
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            updated = conn.execute(
                """UPDATE exported_vouchers SET import_status = ?, import_note = ?
                   WHERE export_id = ? AND number = ? AND book_id = ? AND import_status <> 'imported'
                   AND NOT EXISTS (SELECT 1 FROM voucher_retries r
                                   WHERE r.original_export_id = exported_vouchers.export_id
                                     AND r.original_number = exported_vouchers.number)""",
                (status, note.strip(), export_id, number, book_id),
            )
            if not updated.rowcount:
                raise StorageError("该凭证不存在，或已确认入账或进入修订重导，不能更改原状态。")
