from __future__ import annotations

import json
import os
import sqlite3
from contextlib import closing, contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from importlib.resources import files
from pathlib import Path
from typing import Iterator, Literal

from .models import BookAccount, Issue, ParsedStatement, StatementRow, canonical_json, fingerprint

APPLICATION_ID = 0x4A494E44
MIGRATIONS = ("001_initial.sql", "002_review.sql", "003_review_actions.sql", "004_reconciliation.sql",
              "005_book_setup.sql", "006_batch_book_setup.sql", "007_recalculated_rows.sql")
RowStatus = Literal["new", "invalid", "duplicate", "conflict", "possible_duplicate"]


class StorageError(ValueError):
    """Storage settings or data would violate a batch invariant."""


class StaleBookSetup(StorageError):
    """The book settings have changed since this operation began."""


def _register_account(conn: sqlite3.Connection, account: BookAccount) -> None:
    scope = account.book_id, account.account_id
    conn.execute(
        "INSERT INTO accounts(book_id, id, number) VALUES (?, ?, ?) ON CONFLICT DO NOTHING",
        (*scope, account.number),
    )
    stored = conn.execute("SELECT number FROM accounts WHERE book_id = ? AND id = ?", scope).fetchone()
    if stored is None or stored["number"] != account.number:
        raise StorageError("账户标识与完整账号的原绑定不可更改。停用旧户请从新版本移除，新增账户请用未使用的标识和账号。")


def default_database() -> Path:
    local = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    return local / "Jindie" / "statements.sqlite3"


@dataclass(frozen=True)
class SavedBatch:
    id: int
    reused: bool


def _migration_statements(script: str) -> Iterator[str]:
    pending = ""
    for line in script.splitlines(keepends=True):
        pending += line
        if sqlite3.complete_statement(pending):
            yield pending
            pending = ""
    if pending.strip():
        raise StorageError("数据库迁移脚本不完整，请检查安装文件。")


class StatementStore:
    def __init__(self, database: Path):
        self.database = database

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        with closing(sqlite3.connect(self.database)) as conn:
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON")
            with conn:
                yield conn

    def initialize(self) -> None:
        self.database.parent.mkdir(parents=True, exist_ok=True)
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            app_id = conn.execute("PRAGMA application_id").fetchone()[0]
            version = conn.execute("PRAGMA user_version").fetchone()[0]
            if app_id != APPLICATION_ID and (
                app_id != 0 or version != 0 or conn.execute("SELECT 1 FROM sqlite_schema LIMIT 1").fetchone()
            ):
                raise StorageError("所选文件不是本工具的数据库，请使用其他路径。")
            if version > len(MIGRATIONS):
                raise StorageError("数据库由更高版本创建，请使用对应版本的程序。")
            if version and version < len(MIGRATIONS):
                suffix = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
                backup = self.database.with_name(f"{self.database.name}.v{version}-{suffix}.bak")
                with self._connection() as source, closing(sqlite3.connect(backup)) as target:
                    source.backup(target)
            for index in range(version, len(MIGRATIONS)):
                script = files("jindie").joinpath("migrations", MIGRATIONS[index]).read_text(encoding="utf-8")
                for statement in _migration_statements(script):
                    conn.execute(statement)
                conn.execute(f"PRAGMA user_version = {index + 1}")
            conn.execute(f"PRAGMA application_id = {APPLICATION_ID}")

    def create_book_setup(
        self, book_id: str, accounts: tuple[BookAccount, ...], config: dict, template: bytes,
    ) -> None:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            if not conn.execute("INSERT INTO books(id) VALUES (?) ON CONFLICT DO NOTHING", (book_id,)).rowcount:
                raise StorageError("该账套标识已经存在或有历史流水，请核对后换用新账套标识。")
            self._insert_book_setup(conn, book_id, accounts, config, template)

    @staticmethod
    def _insert_book_setup(
        conn: sqlite3.Connection, book_id: str, accounts: tuple[BookAccount, ...], config: dict, template: bytes,
    ) -> None:
        for account in accounts:
            if account.book_id != book_id:
                raise StorageError("银行账户不属于该账套。")
            _register_account(conn, account)
        conn.execute(
            "INSERT INTO book_setup_versions(book_id, version, config_json, template) VALUES (?, ?, ?, ?)",
            (book_id, config["ledger"]["version"], canonical_json(config), template),
        )

    def publish_book_setup(
        self, book_id: str, accounts: tuple[BookAccount, ...], config: dict, template: bytes, *, previous_version: int,
    ) -> None:
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            current = conn.execute(
                "SELECT MAX(version) FROM book_setup_versions WHERE book_id = ?", (book_id,),
            ).fetchone()[0]
            if current != previous_version:
                raise StaleBookSetup("账套配置已更新，请刷新配置页并基于最新版本重新提交。")
            if config["ledger"]["version"] != previous_version + 1:
                raise StorageError("新账套配置版本号必须紧接现有版本。")
            self._insert_book_setup(conn, book_id, accounts, config, template)

    def current_book_setups(self) -> list[dict]:
        with self._connection() as conn:
            return [json.loads(row["config_json"]) for row in conn.execute(
                """SELECT v.config_json FROM book_setup_versions v
                   JOIN (SELECT book_id, MAX(version) AS version FROM book_setup_versions GROUP BY book_id) active
                     ON active.book_id = v.book_id AND active.version = v.version
                   ORDER BY v.book_id"""
            )]

    def book_setup(self, book_id: str, version: int) -> tuple[dict, bytes]:
        with self._connection() as conn:
            row = conn.execute(
                "SELECT config_json, template FROM book_setup_versions WHERE book_id = ? AND version = ?",
                (book_id, version),
            ).fetchone()
            if row is None:
                raise StorageError("找不到本批使用的账套模板版本，不能生成正式导入文件。")
            return json.loads(row["config_json"]), row["template"]

    def book_setup_config(self, book_id: str, version: int) -> dict:
        with self._connection() as conn:
            row = conn.execute(
                "SELECT config_json FROM book_setup_versions WHERE book_id = ? AND version = ?", (book_id, version),
            ).fetchone()
            if row is None:
                raise StorageError("找不到该账套使用的配置版本。")
            return json.loads(row["config_json"])

    def save(self, statement: ParsedStatement, *, expected_setup: int | None = None) -> SavedBatch:
        self.initialize()
        account = statement.account
        scope = account.book_id, account.account_id
        profile = statement.profile_id, statement.profile_version
        options = canonical_json({
            "confirmed_binding": statement.confirmed_binding,
            "posting_preference": statement.posting_preference,
        })
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("INSERT INTO books(id) VALUES (?) ON CONFLICT DO NOTHING", (account.book_id,))
            setup_version = conn.execute(
                "SELECT MAX(version) FROM book_setup_versions WHERE book_id = ?", (account.book_id,),
            ).fetchone()[0]
            if expected_setup is not None and setup_version != expected_setup:
                raise StaleBookSetup("上传期间账套配置已更新，请刷新上传页并重新选择账户。")
            _register_account(conn, account)
            conn.execute(
                "INSERT INTO bank_profile_versions VALUES (?, ?, ?) ON CONFLICT DO NOTHING", (*profile, statement.profile_json)
            )
            stored_profile = conn.execute(
                "SELECT config_json FROM bank_profile_versions WHERE id = ? AND version = ?", profile
            ).fetchone()
            if stored_profile["config_json"] != statement.profile_json:
                raise StorageError("银行配置已变更，请保存为新版本，不得覆盖历史配置。")
            identity = (*scope, str(statement.period), *profile, statement.file_hash, statement.sheet, options)
            existing = conn.execute(
                """SELECT id FROM import_batches WHERE book_id = ? AND account_id = ? AND period = ?
                   AND profile_id = ? AND profile_version = ? AND file_hash = ? AND sheet = ? AND options_json = ?""",
                identity,
            ).fetchone()
            if existing is not None:
                return SavedBatch(existing["id"], True)
            batch_id = conn.execute(
                """INSERT INTO import_batches (
                       book_id, account_id, period, profile_id, profile_version, file_hash, sheet, options_json,
                       filename, header_row, warnings_json, setup_version
                   ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id""",
                (*identity, statement.filename, statement.header_row,
                 canonical_json([asdict(warning) for warning in statement.warnings]), setup_version),
            ).fetchone()["id"]
            for row in statement.rows:
                self._save_row(conn, batch_id, statement, row)
            return SavedBatch(batch_id, False)

    def _save_row(
        self, conn: sqlite3.Connection, batch_id: int, statement: ParsedStatement, row: StatementRow
    ) -> None:
        status: RowStatus = "invalid"
        problems = list(row.issues)
        data = row.transaction.payload() if row.transaction else None
        transaction_id = None
        if row.transaction is not None and data is not None:
            status, transaction_id, issue = self._match_transaction(conn, statement, row)
            if issue is not None:
                problems.append(issue)
        raw = dict(row.raw)
        conn.execute(
            """INSERT INTO source_rows (
                   batch_id, row_number, content_hash, raw_json, parsed_json, issues_json, match_status, transaction_id
               ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (batch_id, row.row_number, fingerprint(raw), canonical_json(raw), canonical_json(data) if data is not None else None,
             canonical_json([asdict(issue) for issue in problems]), status, transaction_id),
        )

    def _match_transaction(
        self, conn: sqlite3.Connection, statement: ParsedStatement, row: StatementRow
    ) -> tuple[RowStatus, int | None, Issue | None]:
        tx = row.transaction
        assert tx is not None
        scope = statement.account.book_id, statement.account.account_id
        origin = fingerprint([statement.file_hash, statement.sheet, row.row_number])
        payload = tx.payload()
        content = fingerprint({k: v for k, v in payload.items() if k != "voucher_date"})
        candidate = fingerprint(tx.key_fields())
        existing = conn.execute(
            """SELECT id, content_hash FROM bank_transactions
               WHERE book_id = ? AND account_id = ? AND (origin_key = ? OR reference = ?)
               ORDER BY (origin_key = ?) DESC""",
            (*scope, origin, tx.bank_reference, origin),
        ).fetchall()
        if existing:
            if len(existing) != 1 or existing[0]["content_hash"] != content:
                return "conflict", None, Issue("duplicate_conflict", "流水标识与已有记录一致但内容不同，请人工核对。")
            return "duplicate", existing[0]["id"], None
        possible = conn.execute(
            """SELECT 1 FROM bank_transactions
               WHERE book_id = ? AND account_id = ? AND candidate_hash = ?
               AND file_hash <> ? AND (reference IS NULL OR ? IS NULL) LIMIT 1""",
            (*scope, candidate, statement.file_hash, tx.bank_reference),
        ).fetchone()
        if possible:
            return "possible_duplicate", None, Issue("possible_duplicate", "存在日期、金额及对方相近且缺流水号的记录，请人工核对。")
        inserted = conn.execute(
            """INSERT INTO bank_transactions
               (book_id, account_id, reference, origin_key, file_hash, content_hash, candidate_hash, data_json)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING id""",
            (*scope, tx.bank_reference, origin, statement.file_hash, content, candidate, canonical_json(payload)),
        ).fetchone()
        return "new", inserted["id"], None

    def inspect(self, batch_id: int) -> tuple[sqlite3.Row, list[sqlite3.Row]]:
        self.initialize()
        with self._connection() as conn:
            batch = conn.execute(
                """SELECT b.*, json_extract(v.config_json, '$.name') AS profile_name
                   FROM import_batches b JOIN bank_profile_versions v
                     ON v.id = b.profile_id AND v.version = b.profile_version
                   WHERE b.id = ?""", (batch_id,),
            ).fetchone()
            if batch is None:
                raise StorageError("未找到该批次。")
            rows = conn.execute("SELECT * FROM source_rows WHERE batch_id = ? ORDER BY row_number", (batch_id,))
            return batch, rows.fetchall()

    def recent_batches(self, book_id: str) -> list[sqlite3.Row]:
        self.initialize()
        with self._connection() as conn:
            return conn.execute(
                """SELECT id, account_id, period, filename, sheet, created_at FROM import_batches
                   WHERE book_id = ? ORDER BY id DESC""",
                (book_id,),
            ).fetchall()

    def summary(self, batch_id: int) -> dict[str, object]:
        batch, rows = self.inspect(batch_id)
        counts = {"new": 0, "invalid": 0, "duplicate": 0, "conflict": 0, "possible_duplicate": 0}
        totals: dict[str, dict[str, dict[str, int]]] = {}
        dates: list[str] = []
        flagged: list[dict[str, object]] = []
        for row in rows:
            counts[row["match_status"]] += 1
            if row["parsed_json"]:
                payload = json.loads(row["parsed_json"])
                currency = totals.setdefault(payload["currency"], {})
                totals_for_side = currency.setdefault(payload["direction"], {"count": 0, "amount_minor": 0})
                totals_for_side["count"] += 1
                totals_for_side["amount_minor"] += payload["amount_minor"]
                dates.append(payload["transaction_date"])
            if issues := json.loads(row["issues_json"]):
                flagged.append({"row": row["row_number"], "issues": issues})
        return {
            "batch_id": batch_id,
            "profile": batch["profile_id"],
            "profile_name": batch["profile_name"],
            "profile_version": batch["profile_version"],
            "sheet": batch["sheet"],
            "header_row": batch["header_row"],
            "period": batch["period"],
            "rows": len(rows),
            "counts": counts,
            "amounts_incomplete": counts["invalid"] != 0,
            "date_range": [min(dates), max(dates)] if dates else None,
            "totals_by_currency": totals,
            "warnings": json.loads(batch["warnings_json"]),
            "flagged": flagged,
        }
