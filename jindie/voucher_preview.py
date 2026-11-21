from __future__ import annotations

import json
from dataclasses import asdict, replace
from datetime import date, time
from pathlib import Path
from typing import Mapping

from .accounting import (
    AccountingError, Assignment, Auxiliary, Condition, Exclusion, FixedEntry, Ledger, Proposal, Rule, Source, Subject, propose,
)
from .excel_reader import ImportLimits
from .kingdee_export import ExportOptions, export_workbook
from .models import Issue, Period, StatementTransaction
from .storage import StatementStore


def _unique_index(items, key):
    result = {}
    for item in items:
        index = key(item)
        if index in result:
            raise AccountingError(f"配置包含重复编码或标识：{index}")
        result[index] = item
    return result


def _assignment(data: dict) -> Assignment:
    return Assignment(data["subject"], tuple(Auxiliary(**item) for item in data.get("auxiliary", ())))


def parse_config(data: dict) -> tuple[Ledger, tuple[Rule, ...]]:
    try:
        info = data["ledger"]
        subjects = (
            Subject(item["code"], item["name"], tuple(item.get("required_aux", ())), item.get("quantity_required", False))
            for item in info["subjects"]
        )
        ledger = Ledger(
            info["book_id"], info["version"], info["bank_subjects"], _unique_index(subjects, lambda item: item.code),
            frozenset(Auxiliary(**item) for item in info.get("auxiliary", ())),
            {code: _assignment(item) for code, item in info.get("legacy_assignments", {}).items()},
            info.get("finance_verified", False),
        )
        rules = tuple(
            Rule(**{
                **rule, "counterpart": _assignment(rule["counterpart"]) if rule.get("counterpart") is not None else None,
                "fixed_entries": tuple(
                    FixedEntry(_assignment(item), item["side"], item["amount_minor"])
                    for item in rule.get("fixed_entries", ())
                ),
                "bank_references": tuple(rule.get("bank_references", ())),
                "counterpart_names": tuple(rule.get("counterpart_names", ())),
                "conditions": tuple(
                    Condition(item["field"], item["operation"], tuple(item["values"]))
                    for item in rule.get("conditions", ())
                ),
            })
            for rule in data["rules"]
        )
    except (AttributeError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise AccountingError("凭证配置 JSON 无效，请检查账套主数据、科目和规则字段。") from exc
    return ledger, rules


def load_config(path: Path) -> tuple[Ledger, tuple[Rule, ...]]:
    return parse_config(json.loads(path.read_text(encoding="utf-8-sig")))


def transaction_from_json(data: dict) -> StatementTransaction:
    tx = dict(data)
    for key in ("transaction_date", "posting_date", "voucher_date"):
        tx[key] = date.fromisoformat(tx[key]) if tx[key] else None
    tx["transaction_time"] = time.fromisoformat(tx["transaction_time"]) if tx["transaction_time"] else None
    return StatementTransaction(**tx)


def batch_sources(
    store: StatementStore, batch_id: int, ledger: Ledger, *, account_id: str | None = None,
) -> tuple[Period, tuple[Source, ...], tuple[Exclusion, ...]]:
    batch, rows = store.inspect(batch_id)
    if (batch["book_id"] != ledger.book_id or batch["account_id"] not in ledger.bank_subjects
            or account_id not in (None, batch["account_id"])):
        raise AccountingError("配置的账套或账户与导入批次不一致。")
    period = Period.from_text(batch["period"])
    sources: list[Source] = []
    excluded: list[Exclusion] = []
    for row in rows:
        issues = tuple(Issue(**item) for item in json.loads(row["issues_json"]))
        if issues or row["match_status"] != "new":
            if not issues:
                issues = (Issue("duplicate", "流水已关联到先前批次，不再生成新凭证。"),)
            excluded.append(Exclusion(row["row_number"], issues))
            continue
        sources.append(Source(
            row["transaction_id"], batch["book_id"], batch["account_id"], period, batch_id, batch["sheet"],
            row["row_number"], transaction_from_json(json.loads(row["parsed_json"])),
        ))
    return period, tuple(sources), tuple(excluded)


def plan_batch(
    store: StatementStore, batch_id: int, ledger: Ledger, rules: tuple[Rule, ...],
    confirmed_skips: Mapping[str, str] | None = None,
) -> tuple[Period, Proposal]:
    period, sources, excluded = batch_sources(store, batch_id, ledger)
    proposal = propose(sources, ledger, rules, confirmed_skips=confirmed_skips)
    return period, replace(proposal, exclusions=tuple(sorted([*excluded, *proposal.exclusions], key=lambda e: e.row_number)))


def create_preview(
    store: StatementStore, batch_id: int, config: Path, template: Path, output: Path, voucher_word: str,
    first_number: int, *, skip_file: Path | None = None, include_partial: bool = False,
    text_number: bool = False, used_numbers: frozenset[int] = frozenset(),
) -> dict[str, object]:
    if output.suffix.lower() != ".xlsx":
        raise AccountingError("试算预览的输出文件扩展名须为 .xlsx。")
    ledger, rules = load_config(config)
    skips = None
    if skip_file is not None:
        skips = json.loads(skip_file.read_text(encoding="utf-8-sig"))
        if not isinstance(skips, dict) or any(not isinstance(reason, str) for reason in skips.values()):
            raise AccountingError("已确认跳过文件须用 JSON 对象记录流水号和原因。")
    period, proposal = plan_batch(store, batch_id, ledger, rules, skips)
    report: dict[str, object] = {
        "batch_id": batch_id,
        "mode": "preview",
        "vouchers": len(proposal.vouchers),
        "entries": sum(len(voucher.entries) for voucher in proposal.vouchers),
        "skipped": sum(item.skipped for item in proposal.exclusions),
        "exclusions": [asdict(item) for item in proposal.exclusions],
        "workbook_written": False,
    }
    if not proposal.vouchers or proposal.exclusions and not include_partial:
        return report
    options = ExportOptions(period, voucher_word, first_number, text_number=text_number, used_numbers=used_numbers)
    with template.open("rb") as file:
        result = export_workbook(file.read(ImportLimits().file_bytes + 1), proposal.vouchers, ledger, options)
    with output.open("xb") as file:
        file.write(result.data)
    report["workbook_written"] = True
    report["numbered_sources"] = [
        {"number": item.number, "rows": [source.row_number for source in item.voucher.sources]}
        for item in result.numbered
    ]
    return report
