import hashlib
import json
import re
import sqlite3
import unittest
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from jindie.accounting import AccountingError
from jindie.bank_import import parse_statement
from jindie.kingdee_export import HEADERS, IMPORT_SHEET, ExportError, ExportOptions
from jindie.models import BookAccount
from jindie.storage import APPLICATION_ID, StorageError
from jindie.voucher_workflow import StaleReview, VoucherWorkflow
from jindie.web import create_app

from .fixtures import BOC_ACCOUNT, CCB_ACCOUNT, CCB_HEADERS, boc_rows, ccb_row, make_xls, make_xlsx
from .test_import import PERIOD
from .voucher_fixtures import write_example


def field(response, name):
    return re.search(rf'name="{name}" value="([^"]+)"', response.get_data(as_text=True))[1]


class BrowserWorkflowTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        write_example(self.root)
        self.db = self.root / "test.sqlite3"
        self.config_file = self.root / "voucher-config.json"
        self.template = (self.root / "synthetic-template.xlsx").read_bytes()
        self.config = json.loads(self.config_file.read_text(encoding="utf-8"))
        self.config["ledger"]["finance_verified"] = True
        self.config["web"]["verified_template_hash"] = hashlib.sha256(self.template).hexdigest()
        self.save_config()
        self.app = create_app(self.db, [self.config_file])
        self.app.testing = True
        self.client = self.app.test_client()
        self.store = VoucherWorkflow(self.db)

    def save_config(self):
        self.config_file.write_text(json.dumps(self.config, ensure_ascii=False), encoding="utf-8")

    def upload(self):
        home = self.client.get("/")
        self.assertEqual(home.status_code, 200)
        self.assertNotIn(BOC_ACCOUNT, home.get_data(as_text=True))
        data = {"csrf": field(home, "csrf"), "book": "demo", "account": "boc", "period": "2026-11",
                "statement": (BytesIO((self.root / "boc-voucher-synthetic.xls").read_bytes()), "sample.xls")}
        response = self.client.post("/upload", data=data)
        self.assertEqual(response.status_code, 303)
        return response.location

    def test_browser_review_export_download_and_reopen(self):
        url = self.upload()
        self.assertEqual(self.upload(), url)
        batch = self.client.get(url)
        self.assertIn("44 张凭证", batch.get_data(as_text=True))
        csrf, token = field(batch, "csrf"), field(batch, "token")
        draft_url = f"{url}/drafts/1"
        self.assertEqual(self.client.post(draft_url, data={"operation": "confirm", "token": token}).status_code, 403)
        self.assertEqual(self.client.post(draft_url, data={"csrf": csrf, "operation": "confirm", "token": token},
                                          headers={"Origin": "http://example.invalid"}).status_code, 403)
        self.assertEqual(self.client.post(draft_url, data={"csrf": csrf, "operation": "confirm", "token": token},
                                          headers={"Origin": "http://localhost"}).status_code, 303)
        self.assertEqual(self.client.post(draft_url, data={"csrf": csrf, "operation": "confirm", "token": token}).status_code, 409)
        reviewed = self.client.get(url)
        self.assertIn("1 张已复核待导出", reviewed.get_data(as_text=True))
        edited = self.client.post(draft_url, data={
            "csrf": csrf, "operation": "edit", "token": field(reviewed, "token"), "line": "2",
            "summary": "<b>本批修正</b>", "subject": "1122", "aux": "客户=001",
        })
        self.assertEqual(edited.status_code, 303)
        pending = self.client.get(url)
        self.assertIn("0 张已复核待导出", pending.get_data(as_text=True))
        self.assertIn("&lt;b&gt;本批修正&lt;/b&gt;", pending.get_data(as_text=True))
        self.assertEqual(self.client.post(draft_url, data={
            "csrf": csrf, "operation": "confirm", "token": field(pending, "token"),
        }).status_code, 303)
        ready = self.client.get(url)
        params = {"csrf": csrf, "revision": field(ready, "revision"),
                  "request_id": field(ready, "request_id"), "first_number": "101"}
        self.assertEqual(self.client.post(f"{url}/export", data=params).status_code, 400)
        exported = self.client.post(f"{url}/export", data={**params, "partial": "yes"})
        self.assertEqual(exported.status_code, 303)
        retried = self.client.post(f"{url}/export", data={**params, "partial": "yes"})
        self.assertEqual(retried.location, exported.location)
        self.assertEqual(self.client.post(f"{url}/export", data={**params, "partial": "yes", "first_number": "102"}).status_code, 409)
        self.assertEqual(self.client.get(exported.location).status_code, 200)
        response = self.client.get(f"{exported.location}/download")
        self.assertTrue(response.headers["Content-Disposition"].startswith("attachment;"))
        workbook = load_workbook(BytesIO(response.data), read_only=True)
        try:
            sheet = workbook["凭证模版"]
            self.assertEqual((sheet["C2"].value, sheet["F3"].value, sheet["K3"].value), (101, "<b>本批修正</b>", "001"))
        finally:
            workbook.close()
        status = f"{exported.location}/numbers/101"
        self.assertEqual(self.client.post(status, data={"csrf": csrf, "status": "failed"}).status_code, 400)
        self.assertEqual(self.client.post(status, data={"csrf": csrf, "status": "failed", "note": "合成试导失败"}).status_code, 303)
        self.assertIn("合成试导失败", self.client.get(exported.location).get_data(as_text=True))
        self.assertEqual(self.client.post(status, data={"csrf": csrf, "status": "imported"}).status_code, 303)
        self.assertEqual(self.client.post(status, data={"csrf": csrf, "status": "failed", "note": "重置"}).status_code, 400)
        reopened = create_app(self.db, [self.config_file]).test_client()
        self.assertIn("#1", reopened.get("/history?book=demo&account=boc&period=2026-11").get_data(as_text=True))
        self.assertIn("已导出", reopened.get(url).get_data(as_text=True))
        self.assertEqual(reopened.get(f"{exported.location}/download").data, response.data)

    def test_export_is_atomic_and_checks_local_numbers_and_stale_batch(self):
        url = self.upload()
        batch_id = int(url.rsplit("/", 1)[1])
        _, _, drafts, _ = self.store.review_data(batch_id, "demo")
        first, second = drafts[:2]
        self.store.change_draft("demo", batch_id, first.id, first.token)
        options = ExportOptions(PERIOD, "记", 101, verified_template_hash=hashlib.sha256(self.template).hexdigest())
        with self.assertRaises(StaleReview):
            self.store.create_export("demo", batch_id, 0, "stale", self.template, options, allow_partial=True)
        with self.assertRaisesRegex(ExportError, "核验"):
            self.store.create_export("demo", batch_id, 1, "wrong-template", b"bad", options, allow_partial=True)
        self.assertFalse(self.store.review_data(batch_id, "demo")[3])

        def generate(_):
            return self.store.create_export("demo", batch_id, 1, "same-click", self.template, options, allow_partial=True)

        with ThreadPoolExecutor(max_workers=2) as pool:
            self.assertEqual(list(pool.map(generate, (0, 1))), [1, 1])
        self.store.change_draft("demo", batch_id, second.id, second.token)
        with self.assertRaisesRegex(ExportError, "已使用"):
            self.store.create_export("demo", batch_id, 3, "taken", self.template, options, allow_partial=True)
        self.assertEqual(len(self.store.review_data(batch_id, "demo")[3]), 1)

    def test_unverified_batch_cannot_be_exported_by_updating_startup_config(self):
        self.config["ledger"]["finance_verified"] = False
        self.save_config()
        self.client = create_app(self.db, [self.config_file]).test_client()
        url = self.upload()
        self.config["ledger"]["finance_verified"] = True
        self.save_config()
        self.client = create_app(self.db, [self.config_file]).test_client()
        page = self.client.get(url)
        self.assertIn("尚未完成核验", page.get_data(as_text=True))
        self.assertEqual(self.client.post(f"{url}/export", data={"csrf": field(page, "csrf")}).status_code, 400)

    def test_split_skip_undo_and_merge_preserve_sources_and_reset_review(self):
        url = self.upload()
        page = self.client.get(url)
        csrf = field(page, "csrf")
        batch, _, drafts, _ = self.store.review_data(1, "demo")
        fee = next(draft for draft in drafts if [s.row_number for s in draft.voucher.sources] == [47, 48])
        skip_path = f"{url}/skip/47"
        reason = {"csrf": csrf, "reason": "posted", "note": "与第 51 号历史凭证核对"}
        self.assertEqual(self.client.post(skip_path, data={**reason, "revision": batch["revision"]}).status_code, 400)
        split = self.client.post(f"{url}/regroup", data={
            "csrf": csrf, "revision": batch["revision"], "operation": "split", "ids": str(fee.id),
        })
        self.assertEqual(split.status_code, 303)
        batch, _, drafts, _ = self.store.review_data(1, "demo")
        pieces = [item for item in drafts if item.voucher.sources[0].row_number in (47, 48)]
        self.assertEqual([[e.amount_minor for e in item.voucher.entries] for item in pieces], [[900, 900], [90, 90]])
        self.assertNotIn(fee.id, {item.id for item in drafts})
        self.store.change_draft("demo", 1, pieces[0].id, pieces[0].token)
        batch, _, _, _ = self.store.review_data(1, "demo")
        self.assertEqual(self.client.post(skip_path, data={**reason, "revision": batch["revision"]}).status_code, 303)
        batch, _, drafts, _ = self.store.review_data(1, "demo")
        report = self.client.get(url).get_data(as_text=True)
        self.assertIn("1 笔已跳过", report)
        self.assertIn("与第 51 号历史凭证核对", report)
        with self.assertRaisesRegex(StorageError, "跳过"):
            self.store.regroup("demo", 1, batch["revision"], tuple(item.id for item in pieces), split=False)
        self.assertEqual(self.client.post(skip_path, data={
            "csrf": csrf, "revision": batch["revision"], "operation": "undo",
        }).status_code, 303)
        batch, _, drafts, _ = self.store.review_data(1, "demo")
        self.assertIsNone(next(item.voucher.review_digest for item in drafts if item.id == pieces[0].id))
        other_day = next(item for item in drafts if item.voucher.sources[0].transaction.voucher_date.day != 12)
        with self.assertRaisesRegex(AccountingError, "同一账套"):
            self.store.regroup("demo", 1, batch["revision"], (pieces[0].id, other_day.id), split=False)
        merged, = self.store.regroup("demo", 1, batch["revision"], tuple(item.id for item in pieces), split=False)
        _, _, drafts, _ = self.store.review_data(1, "demo")
        result = next(item for item in drafts if item.id == merged).voucher
        self.assertEqual([s.row_number for s in result.sources], [47, 48])
        self.assertIsNone(result.review_digest)

    def test_manual_voucher_uses_source_amount_and_requires_separate_review(self):
        imported = self.store.save(parse_statement(
            "sample.xls", (self.root / "boc-voucher-synthetic.xls").read_bytes(),
            account=BookAccount("demo", "boc", BOC_ACCOUNT), period=PERIOD,
        ))
        url = f"/books/demo/batches/{imported.id}"
        manual_path = f"{url}/manual/57"
        form = self.client.get(manual_path)
        self.assertEqual(form.status_code, 200)
        self.assertIn(manual_path, self.client.get(url).get_data(as_text=True))
        page = form.get_data(as_text=True)
        self.assertIn("第 57 行", page)
        self.assertIn("流水号 000048", page)
        self.assertIn("贷记 100203", page)
        self.assertIn("必填：供应商", page)
        csrf = field(form, "csrf")
        payload = {"csrf": csrf, "revision": field(form, "revision"), "subject": "2202",
                   "summary": "已核实用途", "reason": "核对纸质付款单后处理"}
        missing_aux = self.client.post(manual_path, data=payload)
        self.assertEqual(missing_aux.status_code, 400)
        self.assertIn("缺少「供应商」编码", missing_aux.get_data(as_text=True))
        self.assertIn("已核实用途", missing_aux.get_data(as_text=True))
        result = self.client.post(manual_path, data={**payload, "aux": "供应商=029"})
        self.assertEqual(result.status_code, 303)
        self.assertEqual(self.client.post(manual_path, data={**payload, "aux": "供应商=029"}).status_code, 409)
        self.assertEqual(self.client.get(manual_path).status_code, 400)
        batch, _, drafts, _ = self.store.review_data(1, "demo")
        draft, = [item for item in drafts if item.voucher.sources[0].row_number == 57]
        self.assertEqual([(e.side, e.assignment.subject, e.amount_minor) for e in draft.voucher.entries],
                         [("debit", "2202", 200), ("credit", "100203", 200)])
        self.assertIsNone(draft.voucher.review_digest)
        options = ExportOptions(PERIOD, "记", 101, verified_template_hash=hashlib.sha256(self.template).hexdigest())
        with self.assertRaisesRegex(StorageError, "没有通过校验"):
            self.store.create_export("demo", 1, batch["revision"], "manual-pending", self.template, options, allow_partial=True)
        self.store.change_draft("demo", 1, draft.id, draft.token)
        export_id = self.store.create_export("demo", 1, batch["revision"] + 1, "manual-reviewed", self.template, options,
                                             allow_partial=True)
        export, items = self.store.export_data("demo", export_id)
        self.assertEqual([voucher.sources[0].row_number for _, voucher in items], [57])
        self.assertNotIn(57, [item["row_number"] for item in json.loads(export["exclusions_json"])])
        with closing(sqlite3.connect(self.db)) as conn:
            details = json.loads(conn.execute("SELECT details_json FROM review_actions WHERE kind = 'manual_voucher'").fetchone()[0])
        self.assertEqual((details["reason"], details["prior_issues"][0]["code"]), (payload["reason"], "unmatched"))

    def test_pending_queue_separates_coding_problems_and_paginates(self):
        rule = self.config["rules"][0]
        rule["counterpart"]["auxiliary"][0]["code"] = "999"
        self.save_config()
        self.client = create_app(self.db, [self.config_file]).test_client()
        batch_url = self.upload()
        first = self.client.get("/pending?book=demo&period=2026-11&group=coding").get_data(as_text=True)
        second = self.client.get("/pending?book=demo&period=2026-11&group=coding&page=2").get_data(as_text=True)
        self.assertIn("下一页", first)
        self.assertEqual(first.count("按批次原配置手工制证"), 30)
        self.assertEqual(second.count("按批次原配置手工制证"), 7)
        self.assertIn("不在本账套主数据中", first)
        self.assertNotIn("无已启用的匹配规则", first)
        self.assertIn(f"{batch_url}?exception_row=10#exception-10", first)
        linked = self.client.get(f"{batch_url}?exception_row=10").get_data(as_text=True)
        self.assertIn('class="card disclosure exception-card" open', linked)
        self.assertIn('id="exception-10"', linked)
        self.assertIn("没有符合条件", self.client.get("/pending?period=2026-10").get_data(as_text=True))
        self.assertEqual(self.client.get("/pending?group=invalid").status_code, 400)
        self.assertEqual(self.client.get("/pending?book=unknown").status_code, 404)

    def test_manual_voucher_cannot_bypass_existing_drafts_or_skips(self):
        url = self.upload()
        form = self.client.get(f"{url}/manual/58")
        csrf, revision = field(form, "csrf"), field(form, "revision")
        self.assertEqual(self.client.get(f"{url}/manual/10").status_code, 400)
        self.assertEqual(self.client.post(f"{url}/skip/58", data={
            "csrf": csrf, "revision": revision, "reason": "posted", "note": "历史号 8",
        }).status_code, 303)
        self.assertEqual(self.client.get(f"{url}/manual/58").status_code, 400)

        rows = boc_rows()[:11]
        duplicate, invalid = [list(row) for row in rows[9:]]
        duplicate[10] = "收货款"
        invalid[1] = "金额不明"
        rows[9:] = [tuple(duplicate), tuple(invalid)]
        response = self.client.post("/upload", data={
            "csrf": csrf, "book": "demo", "account": "boc", "period": "2026-11",
            "statement": (BytesIO(make_xls(rows)), "duplicate-and-invalid.xls"),
        })
        self.assertEqual(response.status_code, 303)
        batch = self.client.get(response.location).get_data(as_text=True)
        self.assertIn("不再生成新凭证", batch)
        self.assertIn("不是有效金额", batch)
        self.assertNotIn("/manual/", batch)
        for row_number in (10, 11):
            self.assertEqual(self.client.get(f"{response.location}/manual/{row_number}").status_code, 400)

    def test_retry_requires_non_import_and_exports_only_the_selected_voucher(self):
        url = self.upload()
        csrf = field(self.client.get(url), "csrf")
        _, _, drafts, _ = self.store.review_data(1, "demo")
        first, second, other = drafts[:3]
        for item in (first, second):
            self.store.change_draft("demo", 1, item.id, item.token)
        options = ExportOptions(PERIOD, "记", 101, verified_template_hash=hashlib.sha256(self.template).hexdigest())
        original = self.store.create_export("demo", 1, 2, "original", self.template, options, allow_partial=True)
        filename = f"/books/demo/exports/{original}/download"
        original_bytes = self.client.get(filename).data
        self.store.record_import("demo", original, 101, "imported", "")
        self.store.record_import("demo", original, 102, "failed", "合成导入错误")
        status_path = f"/books/demo/exports/{original}/numbers"
        self.assertEqual(self.client.post(f"{status_path}/101/retry", data={"csrf": csrf, "confirm": "yes"}).status_code, 400)
        self.assertEqual(self.client.post(f"{status_path}/102/retry", data={"csrf": csrf, "confirm": "yes"}).status_code, 400)
        self.store.record_import("demo", original, 102, "not_imported", "精斗云查无此号")
        retry = self.client.post(f"{status_path}/102/retry", data={"csrf": csrf, "confirm": "yes"})
        self.assertEqual(retry.status_code, 303)
        self.assertEqual(self.client.post(f"{status_path}/102/retry", data={"csrf": csrf, "confirm": "yes"}).location, retry.location)
        with self.assertRaisesRegex(StorageError, "不能更改"):
            self.store.record_import("demo", original, 102, "imported", "")
        _, _, drafts, _ = self.store.review_data(1, "demo")
        draft = next(item for item in drafts if item.original_number == 102)
        self.assertEqual(draft.voucher.sources, second.voucher.sources)
        self.assertIsNone(draft.voucher.review_digest)
        self.store.change_draft("demo", 1, draft.id, draft.token)
        self.store.change_draft("demo", 1, other.id, other.token)
        page = self.client.get(retry.location)
        revision = field(page, "revision")
        payload = {"csrf": csrf, "revision": revision, "request_id": field(page, "request_id"),
                   "retry_draft_id": str(draft.id), "first_number": "102"}
        self.assertEqual(self.client.post(f"{url}/export", data=payload).status_code, 400)
        exported = self.client.post(f"{url}/export", data={**payload, "first_number": "201"})
        self.assertEqual(exported.status_code, 303)
        self.assertIn("原凭证号 102", self.client.get(exported.location).get_data(as_text=True))
        retry_id = int(exported.location.rsplit("/", 1)[1])
        retry_entries = self.store.export_data("demo", retry_id)[1]
        self.assertEqual([(row["number"], v.sources) for row, v in retry_entries], [(201, second.voucher.sources)])
        self.assertEqual(self.client.get(filename).data, original_bytes)
        self.assertIn("不要整批重新导入", self.client.get(f"/books/demo/exports/{original}").get_data(as_text=True))
        original_entries = self.store.export_data("demo", original)[1]
        self.assertEqual([row["import_status"] for row, _ in original_entries], ["imported", "not_imported"])

    def test_migration_backups_and_preserves_existing_bank_batches(self):
        v1_database = self.root / "legacy.sqlite3"
        migration = Path(__file__).resolve().parents[1] / "jindie" / "migrations" / "001_initial.sql"
        with closing(sqlite3.connect(v1_database)) as connection, connection:
            connection.executescript(migration.read_text(encoding="utf-8"))
            connection.execute("PRAGMA user_version = 1")
            connection.execute(f"PRAGMA application_id = {APPLICATION_ID}")
            connection.execute("INSERT INTO books(id) VALUES ('legacy')")
        VoucherWorkflow(v1_database).initialize()
        backup, = self.root.glob("legacy.sqlite3.v1-*.bak")
        with closing(sqlite3.connect(backup)) as original, closing(sqlite3.connect(v1_database)) as migrated:
            self.assertEqual(original.execute("PRAGMA user_version").fetchone()[0], 1)
            self.assertEqual(migrated.execute("SELECT id FROM books").fetchone()[0], "legacy")
            self.assertIn("setup_version", {row[1] for row in migrated.execute("PRAGMA table_info(import_batches)")})


class BookSetupTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        self.db = root / "test.sqlite3"
        write_example(root)
        self.template = (root / "synthetic-template.xlsx").read_bytes()
        self.client = create_app(self.db, []).test_client()
        self.store = VoucherWorkflow(self.db)

    def submit(self, book_id, label, *, code="001", verified="yes", template=None):
        response = self.client.get("/configuration/master-workbook")
        try:
            book = load_workbook(BytesIO(response.data))
        finally:
            response.close()
        for row in (("0010", f"{label}银行"), ("0020", f"{label}另一银行"), ("1122", f"{label}应收", "客户")):
            book["科目"].append(row)
        book["辅助核算"].append(("客户", code))
        book["银行账户"].append(("ccb", "合成建行", CCB_ACCOUNT, "0010", "建行", "记账日"))
        book["银行账户"].append(("boc", "合成中行", BOC_ACCOUNT, "0020", "中行"))
        with BytesIO() as master:
            book.save(master)
            book.close()
            return self.client.post("/configuration/new", data={
                "csrf": field(self.client.get("/configuration/new"), "csrf"),
                "book_id": book_id, "label": label, "word": "记", "currency_code": "RMB",
                "verified": verified, "text_number": "yes" if book_id == "book-b" else "",
                "master": (BytesIO(master.getvalue()), "master.xlsx"),
                "template": (BytesIO(self.template if template is None else template), "empty.xlsx"),
            })

    def upload(self, book, account, filename, data):
        return self.client.post("/upload", data={
            "csrf": field(self.client.get("/"), "csrf"), "book": book, "account": account,
            "period": "2026-11", "statement": (BytesIO(data), filename),
        })

    def test_two_books_keep_codes_templates_and_reference_scopes_separate(self):
        self.assertIn("新增账套", self.client.get("/").get_data(as_text=True))
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-b", "乙账套").status_code, 303)
        self.client = create_app(self.db, []).test_client()
        home = self.client.get("/").get_data(as_text=True)
        self.assertIn("甲账套", home)
        self.assertIn("乙账套", home)
        self.assertNotIn(CCB_ACCOUNT, home)

        ccb = make_xlsx([CCB_HEADERS, ccb_row(posting="2026/11/03")])
        boc = make_xls(boc_rows()[:10])
        for book, label in (("book-a", "甲账套"), ("book-b", "乙账套")):
            uploaded = self.upload(book, "ccb", "ccb.xlsx", ccb)
            self.assertEqual(uploaded.status_code, 303, uploaded.get_data(as_text=True))
            url = uploaded.location
            manual_url = f"{url}/manual/2"
            manual = self.client.get(manual_url)
            self.assertEqual(manual.status_code, 200)
            csrf = field(manual, "csrf")
            created = self.client.post(manual_url, data={
                "csrf": csrf, "revision": field(manual, "revision"), "summary": "货款",
                "subject": "1122", "aux": "客户=001", "reason": "合成样本核验",
            })
            self.assertEqual(created.status_code, 303)
            batch_id = int(url.rsplit("/", 1)[1])
            _, _, (draft,), _ = self.store.review_data(batch_id, book)
            self.assertEqual(self.client.post(f"{url}/drafts/{draft.id}", data={
                "csrf": csrf, "token": draft.token, "operation": "confirm",
            }).status_code, 303)
            ready = self.client.get(url)
            exported = self.client.post(f"{url}/export", data={
                "csrf": csrf, "revision": field(ready, "revision"), "request_id": field(ready, "request_id"),
                "first_number": "101",
            })
            self.assertEqual(exported.status_code, 303, exported.get_data(as_text=True))
            download = self.client.get(f"{exported.location}/download")
            workbook = load_workbook(BytesIO(download.data), read_only=True)
            try:
                row2, row3 = list(workbook["凭证模版"].values)[1:]
                self.assertEqual((row2[0], row2[2], row2[6], row2[7], row3[10]),
                                 ("2026-11-03", 101 if book == "book-a" else "101", "0010", f"{label}银行", "001"))
            finally:
                workbook.close()
                download.close()
            other = "book-b" if book == "book-a" else "book-a"
            self.assertEqual(self.client.get(f"{exported.location.replace(book, other)}/download").status_code, 400)
        extra = self.upload("book-a", "boc", "boc.xls", boc)
        self.assertEqual(extra.status_code, 303, extra.get_data(as_text=True))
        self.assertEqual(self.store.summary(int(extra.location.rsplit("/", 1)[1]))["counts"]["new"], 1)

    def test_rejects_unverified_inputs_without_creating_book(self):
        for options, message in (
            ({"verified": ""}, "勾选财务核验"),
            ({"code": 1}, "须为文本"),
            ({"template": make_xlsx([["不正确的模板"]], "凭证模版")}, "25 列"),
        ):
            with self.subTest(options=options.keys()):
                response = self.submit("book-a", "甲账套", **options)
                self.assertEqual(response.status_code, 400)
                self.assertIn(message, response.get_data(as_text=True))
                self.assertNotIn("甲账套", self.client.get("/").get_data(as_text=True))
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-a", "不能覆盖").status_code, 400)
        self.assertNotIn("不能覆盖", self.client.get("/").get_data(as_text=True))

    def edit_book(self, form, *, code="002", account_number=CCB_ACCOUNT, template=None):
        response = self.client.get("/configuration/books/book-a/master-workbook")
        try:
            workbook = load_workbook(BytesIO(response.data))
        finally:
            response.close()
        accounts = workbook["银行账户"]
        self.assertEqual(accounts["C2"].value, CCB_ACCOUNT)
        accounts["C2"] = account_number
        if accounts.max_row > 2:
            accounts.delete_rows(3)
        workbook["科目"]["B2"] = "银行科目新版名称"
        workbook["辅助核算"]["B2"] = code
        data = BytesIO()
        workbook.save(data)
        workbook.close()
        fields = {
            "csrf": field(form, "csrf"), "version": field(form, "version"), "label": "甲账套",
            "word": "转", "currency_code": "RMB", "text_number": "yes", "verified": "yes",
            "master": (BytesIO(data.getvalue()), "master.xlsx"),
        }
        if template is not None:
            fields["template"] = (BytesIO(template), "new.xlsx")
        return self.client.post("/configuration/books/book-a/edit", data=fields)

    def test_publishing_keeps_unprepared_legacy_batch_on_its_original_version(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        old_data = make_xlsx([CCB_HEADERS, ccb_row()])
        old_statement = parse_statement(
            "before.xlsx", old_data, account=BookAccount("book-a", "ccb", CCB_ACCOUNT),
            period=PERIOD, posting_preference="posting",
        )
        old_id = self.store.save(old_statement).id
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("DROP TABLE recalculated_rows")
            conn.execute("ALTER TABLE import_batches DROP COLUMN setup_version")
            conn.execute("PRAGMA user_version = 5")
        self.client = create_app(self.db, []).test_client()
        form = self.client.get("/configuration/books/book-a/edit")
        # Distinct valid template bytes let us check the exact template selected for each export.
        new_template = make_xlsx([HEADERS], IMPORT_SHEET)
        published = self.edit_book(form, template=new_template)
        self.assertEqual(published.status_code, 303, published.get_data(as_text=True))
        self.assertEqual(self.edit_book(form).status_code, 409)
        self.assertNotIn("合成中行", self.client.get("/").get_data(as_text=True))
        old_url = f"/books/book-a/batches/{old_id}"
        self.assertEqual(self.upload("book-a", "ccb", old_statement.filename, old_data).location, old_url)
        new_response = self.upload("book-a", "ccb", "after.xlsx",
                                   make_xlsx([CCB_HEADERS, ccb_row(reference="000002")]))
        self.assertEqual(new_response.status_code, 303)
        for url, version, code, word, bank, template in (
            (old_url, 1, "001", "记", "甲账套银行", self.template),
            (new_response.location, 2, "002", "转", "银行科目新版名称", new_template),
        ):
            page = self.client.get(url).get_data(as_text=True)
            self.assertIn(f"账套配置 v{version}", page)
            manual = self.client.get(f"{url}/manual/2")
            csrf = field(manual, "csrf")
            created = self.client.post(f"{url}/manual/2", data={
                "csrf": csrf, "revision": field(manual, "revision"), "summary": "货款", "subject": "1122",
                "aux": f"客户={code}", "reason": "按对应版本核对",
            })
            self.assertEqual(created.status_code, 303)
            _, _, (draft,), _ = self.store.review_data(int(url.rsplit("/", 1)[1]), "book-a")
            self.assertEqual(self.client.post(f"{url}/drafts/{draft.id}", data={
                "csrf": csrf, "token": draft.token, "operation": "confirm",
            }).status_code, 303)
            ready = self.client.get(url)
            exported = self.client.post(f"{url}/export", data={
                "csrf": csrf, "revision": field(ready, "revision"), "request_id": field(ready, "request_id"),
                "first_number": "101",
            })
            self.assertEqual(exported.status_code, 303, exported.get_data(as_text=True))
            export_id = int(exported.location.rsplit("/", 1)[1])
            with closing(sqlite3.connect(self.db)) as conn:
                template_hash, = conn.execute(
                    "SELECT template_hash FROM voucher_exports WHERE id = ? AND book_id = 'book-a'", (export_id,),
                ).fetchone()
            self.assertEqual(template_hash, hashlib.sha256(template).hexdigest())
            workbook = load_workbook(BytesIO(self.store.export_file("book-a", export_id)), read_only=True)
            try:
                rows = list(workbook[IMPORT_SHEET].values)
                self.assertEqual((rows[1][1], rows[1][2], rows[1][7], rows[2][10]),
                                 (word, 101 if version == 1 else "101", bank, code))
            finally:
                workbook.close()

    def test_existing_bank_identity_cannot_be_reassigned(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        form = self.client.get("/configuration/books/book-a/edit")
        result = self.edit_book(form, account_number="0000000000000307")
        self.assertEqual(result.status_code, 400)
        self.assertIn("原绑定不可更改", result.get_data(as_text=True))
        self.assertEqual(field(self.client.get("/configuration/books/book-a/edit"), "version"), "1")
        self.assertEqual(self.edit_book(form).status_code, 303)
        _, kept_template = self.store.book_setup("book-a", 2)
        self.assertEqual(kept_template, self.template)

    def preview_bank_format(self, url, rows, *, amount_mode="signed", extra=None):
        form = self.client.get(url)
        return self.client.post(url, data={
            "csrf": field(form, "csrf"), "version": field(form, "version"), "operation": "preview",
            "name": "合成新格式", "amount_mode": amount_mode, "period": "2026-11",
            "own_account": "结算账号", "transaction_date": "发生日", "reference": "明细号",
            "currency": "交易币种", "amount": "本次金额", "income": "流入", "expense": "流出",
            "counterparty_name": "对方名称", "summary": "说明", "sample": (BytesIO(make_xlsx(rows)), "sample.xlsx"),
            **(extra or {}),
        })

    def test_custom_mapping_preview_then_publish_to_one_account(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-b", "乙账套").status_code, 303)
        headers = ("发生日", "明细号", "本次金额", "结算账号", "对方名称", "交易币种", "说明")
        rows = [headers, ("2026-11-02", "000001", 15.20, CCB_ACCOUNT, "测试对方", "CNY", "试收"),
                ("2026-11-03", "000002", -2, CCB_ACCOUNT, "测试对方", "RMB", "试支")]
        before = self.upload("book-a", "ccb", "old.xlsx", make_xlsx([CCB_HEADERS, ccb_row(reference="000003")]))
        url = "/configuration/books/book-a/accounts/ccb/format"
        preview = self.preview_bank_format(url, rows)
        self.assertEqual(preview.status_code, 200)
        page = preview.get_data(as_text=True)
        self.assertIn("收入 1 笔，15.20 元；支出 1 笔，2.00 元；异常 0 笔", page)
        self.assertIn("000001", page)
        self.assertEqual(len(self.store.recent_batches("book-a")), 1)
        self.assertEqual(self.store.current_book_setups()[0]["ledger"]["version"], 1)
        data = {"csrf": field(preview, "csrf"), "operation": "publish", "preview_token": field(preview, "preview_token")}
        self.assertEqual(self.client.post(url, data=data).status_code, 400)
        data["verified"] = "yes"
        self.assertEqual(self.client.post(url.replace("book-a", "book-b"), data=data).status_code, 400)
        self.assertEqual(self.client.post(url.replace("/ccb/", "/boc/"), data=data).status_code, 400)
        self.assertEqual(self.client.post(url, data=data).status_code, 303)
        self.assertEqual(self.client.post(url, data=data).status_code, 409)
        self.assertEqual(len(self.store.recent_batches("book-a")), 1)
        self.assertIn("账套配置 v1", self.client.get(before.location).get_data(as_text=True))

        form = self.client.get("/configuration/books/book-a/edit")
        revised = self.edit_book(form)
        self.assertEqual(revised.status_code, 303, revised.get_data(as_text=True))
        self.client = create_app(self.db, []).test_client()
        workbook = make_xlsx(rows)
        self.assertEqual(self.upload("book-b", "ccb", "custom.xlsx", workbook).status_code, 400)
        imported = self.upload("book-a", "ccb", "custom.xlsx", workbook)
        self.assertEqual(imported.status_code, 303, imported.get_data(as_text=True))
        batch = self.client.get(imported.location).get_data(as_text=True)
        self.assertIn("账套配置 v3", batch)
        self.assertIn("合成新格式", batch)
        summary = self.store.summary(int(imported.location.rsplit("/", 1)[1]))
        self.assertEqual(summary["counts"]["new"], 2)

    def test_split_mapping_does_not_offer_publication_for_wrong_account_or_duplicate_reference(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        url = "/configuration/books/book-a/accounts/ccb/format"
        headings = ("发生日", "明细号", "结算账号", "流入", "流出")
        income = ("2026-11-02", "000001", CCB_ACCOUNT, 10, None)
        spend = ("2026-11-03", "000002", CCB_ACCOUNT, None, 3)
        for invalid, message in (
            ((*spend[:2], BOC_ACCOUNT, *spend[3:]), "本方账号与所选账户不一致"),
            ((spend[0], income[1], *spend[2:]), "重复银行流水号"),
        ):
            with self.subTest(message=message):
                preview = self.preview_bank_format(
                    url, [headings, income, invalid], amount_mode="split",
                    extra={"currency": "", "cny_only": "yes", "summary": "", "counterparty_name": ""},
                )
                self.assertEqual(preview.status_code, 200)
                self.assertIn(message, preview.get_data(as_text=True))
                self.assertNotIn('name="preview_token"', preview.get_data(as_text=True))
        valid = self.preview_bank_format(
            url, [headings, income, spend], amount_mode="split",
            extra={"currency": "", "cny_only": "yes", "summary": "", "counterparty_name": ""},
        )
        self.assertIn('name="preview_token"', valid.get_data(as_text=True))
        self.assertEqual(self.store.recent_batches("book-a"), [])

    def sample_for_rules(self, book, first_reference, second_reference):
        response = self.upload(book, "ccb", "rules.xlsx", make_xlsx([
            CCB_HEADERS, ccb_row(reference=first_reference, remark="货款"),
            ccb_row(reference=second_reference, remark="保证金"),
        ]))
        self.assertEqual(response.status_code, 303)
        return response.location

    def preview_rule(self, url, batch_url, **changes):
        form = self.client.get(url)
        return self.client.post(url, data={
            "operation": "preview", "csrf": field(form, "csrf"), "version": changes.get("version", "1"),
            "rule_id": "receive-goods", "account_id": "ccb", "direction": "income", "party": "合成对方",
            "field": "remark", "operation_type": "equals", "text": "货款", "subject": "1122", "aux": "客户=001",
            "summary": "收货款-{counterparty}", "batch_id": batch_url.rsplit("/", 1)[1], **changes,
        })

    def publish_rule(self, url, preview, *, verified="yes"):
        return self.client.post(url, data={
            "operation": "publish", "csrf": field(preview, "csrf"), "preview_token": field(preview, "preview_token"),
            "verified": verified,
        })

    def test_published_rule_creates_only_matching_unreviewed_drafts_in_new_batches(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-b", "乙账套").status_code, 303)
        url = "/configuration/books/book-a/rules"
        original = self.sample_for_rules("book-a", "000001", "000002")
        preview = self.preview_rule(url, original)
        self.assertEqual(preview.status_code, 200)
        self.assertIn("命中 1 笔，能生成 1 张草稿", preview.get_data(as_text=True))
        self.assertIn("收货款-合成对方", preview.get_data(as_text=True))
        self.assertEqual(self.store.current_book_setups()[0]["ledger"]["version"], 1)
        self.assertEqual(self.publish_rule(url, preview, verified="").status_code, 400)
        self.assertEqual(self.publish_rule(url.replace("book-a", "book-b"), preview).status_code, 400)
        self.assertEqual(self.publish_rule(url, preview).status_code, 303)
        self.assertEqual(self.publish_rule(url, preview).status_code, 409)
        self.assertIn("0 张凭证", self.client.get(original).get_data(as_text=True))

        next_batch = self.sample_for_rules("book-a", "000003", "000004")
        batch_id = int(next_batch.rsplit("/", 1)[1])
        _, ledger, (draft,), _ = self.store.review_data(batch_id, "book-a")
        self.assertEqual((ledger.version, draft.voucher.review_digest, draft.voucher.rule_versions),
                         (2, None, (("receive-goods", 1),)))
        self.assertEqual([source.row_number for source in draft.voucher.sources], [2])
        self.assertEqual([(entry.assignment.subject, entry.side, entry.amount_minor) for entry in draft.voucher.entries],
                         [("0010", "debit", 1000), ("1122", "credit", 1000)])
        self.assertEqual([(aux.category, aux.code) for aux in draft.voucher.entries[1].assignment.auxiliary],
                         [("客户", "001")])
        self.assertIn("0 张已复核待导出", self.client.get(next_batch).get_data(as_text=True))

        edit = self.preview_rule(url, original, version="2", editing="receive-goods", summary="新版-{counterparty}")
        self.assertEqual(self.publish_rule(url, edit).status_code, 303)
        form = self.client.get(url)
        self.assertEqual(self.client.post(url, data={
            "csrf": field(form, "csrf"), "version": "3", "operation": "disable", "rule_id": "receive-goods",
            "verified": "yes",
        }).status_code, 303)
        after_disable = self.sample_for_rules("book-a", "000005", "000006")
        self.assertIn("0 张凭证", self.client.get(after_disable).get_data(as_text=True))
        old_voucher = self.store.review_data(batch_id, "book-a")[2][0].voucher
        self.assertEqual(old_voucher.entries[0].summary, "收货款-合成对方")

    def test_rule_trial_blocks_wrong_scope_bad_codes_no_hits_and_conflicts(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-b", "乙账套").status_code, 303)
        url = "/configuration/books/book-a/rules"
        sample = self.sample_for_rules("book-a", "000001", "000002")
        other = self.sample_for_rules("book-b", "000001", "000002")
        for source, options, error in (
            (other, {}, "账套或账户与导入批次不一致"),
            (sample, {"account_id": "boc"}, "账套或账户与导入批次不一致"),
            (sample, {"aux": "客户=999"}, "组合不在本账套主数据中"),
            (sample, {"party": "", "field": "", "text": ""}, "不能仅按收支方向匹配"),
        ):
            response = self.preview_rule(url, source, **options)
            self.assertEqual(response.status_code, 400)
            self.assertIn(error, response.get_data(as_text=True))
        missed = self.preview_rule(url, sample, party="不存在的对方")
        self.assertNotIn('name="preview_token"', missed.get_data(as_text=True))
        self.assertEqual(self.publish_rule(url, self.preview_rule(url, sample)).status_code, 303)
        conflict = self.preview_rule(url, sample, version="2", rule_id="conflict", summary="不同摘要")
        self.assertIn("同优先级规则结果不同", conflict.get_data(as_text=True))
        self.assertNotIn('name="preview_token"', conflict.get_data(as_text=True))

    def test_pending_queue_tracks_manual_resolution_and_skip_undo_without_crossing_books(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        self.assertEqual(self.submit("book-b", "乙账套").status_code, 303)
        book_a = self.sample_for_rules("book-a", "000001", "000002")
        book_b = self.sample_for_rules("book-b", "000001", "000002")
        queue = "/pending?book=book-a&group=unmatched"
        manual_two, manual_three = f"{book_a}/manual/2", f"{book_a}/manual/3"
        page = self.client.get(queue).get_data(as_text=True)
        self.assertIn(manual_two, page)
        self.assertIn(manual_three, page)
        self.assertNotIn(book_b, page)
        self.assertIn(book_b, self.client.get("/pending").get_data(as_text=True))

        form = self.client.get(manual_two)
        csrf = field(form, "csrf")
        response = self.client.post(manual_two, data={
            "csrf": csrf, "revision": field(form, "revision"), "summary": "核对货款",
            "subject": "1122", "aux": "客户=001", "reason": "手工核对账套凭证依据",
        })
        self.assertEqual(response.status_code, 303)
        page = self.client.get(queue).get_data(as_text=True)
        self.assertNotIn(manual_two, page)
        self.assertIn(manual_three, page)
        revision = field(self.client.get(manual_three), "revision")
        skip = f"{book_a}/skip/3"
        self.assertEqual(self.client.post(skip, data={
            "csrf": csrf, "revision": revision, "reason": "posted", "note": "已核对入账凭证",
        }).status_code, 303)
        self.assertIn("没有符合条件", self.client.get(queue).get_data(as_text=True))
        batch, _, (draft,), _ = self.store.review_data(int(book_a.rsplit("/", 1)[1]), "book-a")
        self.assertIsNone(draft.voucher.review_digest)
        self.assertEqual(self.client.post(skip, data={
            "csrf": csrf, "revision": batch["revision"], "operation": "undo",
        }).status_code, 303)
        self.assertIn(manual_three, self.client.get(queue).get_data(as_text=True))

    def test_recalculate_only_selected_row_and_preserve_original_export(self):
        self.assertEqual(self.submit("book-a", "甲账套").status_code, 303)
        rules = "/configuration/books/book-a/rules"
        sample = self.sample_for_rules("book-a", "000001", "000002")
        self.assertEqual(self.publish_rule(rules, self.preview_rule(rules, sample)).status_code, 303)
        data = make_xlsx([CCB_HEADERS, ccb_row(reference="000003", remark="货款"),
                          ccb_row(reference="000004", remark="保证金"), ccb_row(reference="000005", remark="其他")])
        original = self.upload("book-a", "ccb", "mixed.xlsx", data).location
        original_id = int(original.rsplit("/", 1)[1])

        def review_and_export(url, number, *, partial=False):
            batch_id = int(url.rsplit("/", 1)[1])
            _, _, (draft,), _ = self.store.review_data(batch_id, "book-a")
            self.assertIsNone(draft.voucher.review_digest)
            page = self.client.get(url)
            csrf = field(page, "csrf")
            self.assertEqual(self.client.post(f"{url}/drafts/{draft.id}", data={
                "csrf": csrf, "operation": "confirm", "token": draft.token,
            }).status_code, 303)
            ready = self.client.get(url)
            response = self.client.post(f"{url}/export", data={
                "csrf": csrf, "revision": field(ready, "revision"), "request_id": field(ready, "request_id"),
                "first_number": number, "partial": "yes" if partial else "",
            })
            self.assertEqual(response.status_code, 303, response.get_data(as_text=True))
            return int(response.location.rsplit("/", 1)[1])

        original_export = review_and_export(original, "101", partial=True)
        original_bytes = self.store.export_file("book-a", original_export)
        recalculate = f"{original}/recalculate/3"
        self.assertNotIn(recalculate, self.client.get("/pending").get_data(as_text=True))
        self.assertEqual(self.client.get(recalculate).status_code, 400)
        rule = self.preview_rule(rules, original, version="2", rule_id="deposit", text="保证金", summary="保证金")
        self.assertEqual(self.publish_rule(rules, rule).status_code, 303)
        self.assertIn(recalculate, self.client.get("/pending").get_data(as_text=True))
        failed = self.client.get(f"{original}/recalculate/4")
        self.assertIn("仍不能生成有效草稿", failed.get_data(as_text=True))
        self.assertNotIn('name="preview_token"', failed.get_data(as_text=True))
        trial = self.client.get(recalculate)
        self.assertIn("原配置 v2 → 最新配置 v3", trial.get_data(as_text=True))
        self.assertIn("deposit v1", trial.get_data(as_text=True))
        fields = {"csrf": field(trial, "csrf"), "preview_token": field(trial, "preview_token"), "verified": "yes"}
        self.assertEqual(self.client.post(f"{original}/recalculate/4", data=fields).status_code, 400)
        self.assertEqual(self.client.post(recalculate, data={**fields, "verified": ""}).status_code, 400)
        self.assertEqual(len(self.store.recent_batches("book-a")), 2)

        self.assertEqual(self.client.post(rules, data={
            "csrf": fields["csrf"], "version": "3", "operation": "disable", "rule_id": "receive-goods", "verified": "yes",
        }).status_code, 303)
        self.assertEqual(self.client.post(recalculate, data=fields).status_code, 409)
        fields["preview_token"] = field(self.client.get(recalculate), "preview_token")
        transferred = self.client.post(recalculate, data=fields)
        self.assertEqual(transferred.status_code, 303, transferred.get_data(as_text=True))
        self.assertEqual(self.client.post(recalculate, data=fields).status_code, 409)
        self.assertEqual(len(self.store.recent_batches("book-a")), 3)
        target = transferred.location.split("?", 1)[0]
        self.assertIn(f"批次 #{original_id} 第 3 行", self.client.get(target).get_data(as_text=True))
        self.assertIn("打开重算批次", self.client.get(f"{original}?exception_row=3").get_data(as_text=True))
        page = self.client.get("/pending").get_data(as_text=True)
        self.assertNotIn(recalculate, page)
        self.assertIn(f"{original}/recalculate/4", page)
        self.assertEqual(self.client.get(f"{original}/manual/3").status_code, 400)
        old_batch, old_ledger, old_drafts, _ = self.store.review_data(original_id, "book-a")
        self.assertEqual((old_ledger.version, old_drafts[0].export_id), (2, original_export))
        self.assertEqual(self.client.post(f"{original}/skip/3", data={
            "csrf": fields["csrf"], "revision": old_batch["revision"], "reason": "posted", "note": "测试",
        }).status_code, 400)
        _, ledger, (draft,), _ = self.store.review_data(int(target.rsplit("/", 1)[1]), "book-a")
        self.assertEqual((ledger.version, draft.voucher.sources[0].row_number), (4, 3))
        self.assertEqual(draft.voucher.rule_versions, (("deposit", 1),))
        self.assertEqual(self.upload("book-a", "ccb", "mixed.xlsx", data).location, original)
        new_export = review_and_export(target, "102")
        self.assertNotEqual(original_export, new_export)
        self.assertEqual(self.store.export_file("book-a", original_export), original_bytes)
        original_info, original_rows = self.store.inspect(original_id)
        target_info, target_rows = self.store.inspect(int(target.rsplit("/", 1)[1]))
        self.assertEqual(original_rows[1]["transaction_id"], target_rows[0]["transaction_id"])
        self.assertEqual((original_info["file_hash"], original_info["sheet"]),
                         (target_info["file_hash"], target_info["sheet"]))
