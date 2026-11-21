import sqlite3
from contextlib import closing
import unittest
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory

from jindie.bank_profiles import CCB_0306
from jindie.storage import StatementStore, StorageError

from .fixtures import ccb_row
from .test_import import CCB, parse_ccb


class StatementStoreTests(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.database = Path(self.directory.name) / "test.sqlite3"
        self.store = StatementStore(self.database)

    def test_reupload_restores_batch_without_duplicate_and_flags_invalid_totals(self):
        parsed = parse_ccb([ccb_row(), ccb_row(income="bad", reference="000002")])
        with ThreadPoolExecutor(max_workers=2) as pool:
            first, second = list(pool.map(self.store.save, [parsed, parsed]))
        self.assertEqual(first.id, second.id)
        self.assertEqual(sorted([first.reused, second.reused]), [False, True])
        report = StatementStore(self.database).summary(first.id)
        self.assertEqual(report["counts"]["new"], 1)
        self.assertEqual(report["counts"]["invalid"], 1)
        self.assertTrue(report["amounts_incomplete"])
        self.assertEqual(report["flagged"][0]["row"], 3)
        self.assertEqual(report["totals_by_currency"]["CNY"]["income"]["amount_minor"], 1000)

    def test_reference_dedupe_ignores_row_order_but_flags_conflicting_amounts(self):
        row_1, row_2 = ccb_row(), ccb_row(reference="000002")
        original_id = self.store.save(parse_ccb([row_1, row_2])).id
        reversed_id = self.store.save(parse_ccb([row_2, row_1])).id
        _, original = self.store.inspect(original_id)
        _, reversed_rows = self.store.inspect(reversed_id)
        self.assertEqual([r["match_status"] for r in reversed_rows], ["duplicate", "duplicate"])
        self.assertEqual([r["transaction_id"] for r in reversed_rows], [r["transaction_id"] for r in reversed(original)])
        conflict = self.store.save(parse_ccb([ccb_row(income=30)])).id
        self.assertEqual(self.store.summary(conflict)["counts"]["conflict"], 1)

    def test_missing_reference_only_flags_duplicates_across_files(self):
        row = ccb_row(reference=None)
        first = self.store.save(parse_ccb([row, row]))
        _, rows = self.store.inspect(first.id)
        self.assertEqual([r["match_status"] for r in rows], ["new", "new"])
        self.assertNotEqual(rows[0]["transaction_id"], rows[1]["transaction_id"])
        duplicate = self.store.save(parse_ccb([row, ccb_row(reference="different", income=20)]))
        self.assertEqual(self.store.summary(duplicate.id)["counts"]["possible_duplicate"], 1)

    def test_book_isolation_and_historical_profile_versions(self):
        original = parse_ccb([ccb_row()])
        self.store.save(original)
        other = self.store.save(replace(original, account=replace(CCB, book_id="another-book")))
        self.assertEqual(self.store.summary(other.id)["counts"]["new"], 1)
        changed = parse_ccb([ccb_row()], profiles=(replace(CCB_0306, name="更新的银行配置"),))
        with self.assertRaisesRegex(StorageError, "新版本"):
            self.store.save(changed)
        next_version = CCB_0306.version + 1
        versioned = parse_ccb([ccb_row()], profiles=(replace(CCB_0306, version=next_version, name="更新的银行配置"),))
        next_batch = self.store.save(versioned)
        self.assertEqual(self.store.summary(next_batch.id)["profile_version"], next_version)

    def test_database_migration_never_reuses_unrelated_database(self):
        with closing(sqlite3.connect(self.database)) as conn, conn:
            conn.execute("CREATE TABLE unrelated (id INTEGER)")
        with self.assertRaisesRegex(StorageError, "不是本工具"):
            self.store.initialize()
