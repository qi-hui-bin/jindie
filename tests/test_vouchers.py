import hashlib
import json
import unittest
from dataclasses import replace
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from jindie.accounting import (
    Assignment, Auxiliary, Condition, Ledger, Rule, Source, Subject, edit_entries, propose, review,
)
from jindie.bank_import import parse_statement
from jindie.kingdee_export import ExportError, ExportOptions, export_workbook
from jindie.models import BookAccount
from jindie.storage import StatementStore
from jindie.voucher_preview import create_preview, load_config, plan_batch

from .fixtures import BOC_ACCOUNT, ccb_row, make_xlsx
from .test_import import PERIOD, parse_ccb
from .voucher_fixtures import write_example

EXPECTED_HEADERS = (
    "日期", "凭证字", "凭证号", "附件数", "分录序号", "摘要", "科目代码", "科目名称", "借方金额", "贷方金额",
    "客户", "供应商", "职员", "项目", "部门", "存货", "自定义辅助核算类别", "自定义辅助核算编码",
    "自定义辅助核算类别1", "自定义辅助核算编码1", "数量", "单价", "原币金额", "币别", "汇率",
)


class VoucherTests(unittest.TestCase):
    def test_rules_keep_same_payee_businesses_separate_and_show_conflicts(self):
        txs = parse_ccb([
            ccb_row(income=0, expense=8, remark="代采"),
            ccb_row(income=0, expense=9, remark="报销", reference="000002"),
            ccb_row(income=0, expense=10, remark="薪资", reference="000003"),
        ]).rows
        sources = tuple(Source(i, "ccb", "0306", PERIOD, 1, "流水", row.row_number, row.transaction)
                        for i, row in enumerate(txs, 1))
        project = Auxiliary("项目", "029")
        ledger = Ledger("ccb", 1, {"0306": "1002001"}, {
            "1002001": Subject("1002001", "银行存款"),
            "1221001": Subject("1221001", "代采款", ("项目",)),
            "2241001": Subject("2241001", "报销款", ("项目",)),
        }, frozenset({project}))
        rules = tuple(
            Rule(id=business, version=1, book_id="ccb", direction="expense", counterpart=Assignment(code, (project,)),
                 summary="{remark}-{counterparty}", account_id="0306", counterpart_names=("合成对方",),
                 conditions=(Condition("remark", "equals", (business,)),), enabled=True)
            for business, code in (("代采", "1221001"), ("报销", "2241001"))
        )
        wrong_book = replace(rules[0], book_id="boc", priority=99, counterpart=rules[1].counterpart)
        fallback = replace(rules[1], id="fallback", conditions=(), counterpart_names=(), enabled=False)
        proposal = propose(sources, ledger, (*rules, wrong_book, fallback))
        self.assertEqual([v.entries[0].assignment.subject for v in proposal.vouchers], ["1221001", "2241001"])
        self.assertEqual(proposal.exclusions[0].issues[0].code, "unmatched")
        ambiguous = replace(rules[0], id="conflict", counterpart=rules[1].counterpart)
        conflict = propose(sources, ledger, (*rules, ambiguous))
        self.assertEqual(conflict.exclusions[0].issues[0].code, "rule_conflict")
        invalid = replace(rules[0], counterpart=Assignment("1221001"))
        self.assertEqual(propose(sources, ledger, (invalid,)).exclusions[0].issues[0].code, "missing_auxiliary")

    def test_synthetic_boc_preview_has_44_vouchers_91_entries_and_exact_fee_rows(self):
        with TemporaryDirectory() as root:
            directory = Path(root)
            write_example(directory)
            store = StatementStore(directory / "statements.sqlite3")
            filename = "boc-voucher-synthetic.xls"
            statement = parse_statement(
                filename, (directory / filename).read_bytes(), account=BookAccount("demo", "boc", BOC_ACCOUNT), period=PERIOD,
            )
            batch_id = store.save(statement).id
            config, skips = directory / "voucher-config.json", directory / "confirmed-skips.json"
            template = directory / "synthetic-template.xlsx"
            output = directory / "preview.xlsx"
            arguments = (store, batch_id, config, template, output, "记", 101)
            blocked = create_preview(*arguments)
            self.assertFalse(blocked["workbook_written"])
            self.assertFalse(output.exists())
            self.assertEqual([item["issues"][0]["code"] for item in blocked["exclusions"]], ["unmatched", "unmatched"])
            report = create_preview(*arguments, skip_file=skips, include_partial=True)
            self.assertEqual((report["vouchers"], report["entries"], report["skipped"]), (44, 91, 2))
            self.assertEqual([item["row_number"] for item in report["exclusions"]], [57, 58])
            self.assertIn([47, 48], [group["rows"] for group in report["numbered_sources"]])
            workbook = load_workbook(output, read_only=True)
            try:
                self.assertNotIn("凭证模版", workbook.sheetnames)
                self.assertEqual(workbook["合成备查"]["A1"].value, "用于测试的备查页")
                header, *entries = list(workbook["试算预览（不可导入）"].values)
                self.assertEqual(header, EXPECTED_HEADERS)
                self.assertEqual(len(entries), 91)
                self.assertEqual(sorted({row[2] for row in entries}), list(range(101, 145)))
                fee_number = next(row[2] for row in entries if row[0] == "2026-11-12" and row[6] == "560303")
                fee_rows = [row for row in entries if row[2] == fee_number]
                self.assertEqual([(r[4], r[6], r[8], r[9], r[22:25]) for r in fee_rows], [
                    (1, "560303", 9.9, None, (9.9, "RMB", 1)),
                    (2, "100203", None, 9, (9, "RMB", 1)),
                    (3, "100203", None, 0.9, (0.9, "RMB", 1)),
                ])
                self.assertTrue(any(row[6] == "2202" and row[11] == "029" for row in entries))
            finally:
                workbook.close()

            ledger, rules = load_config(config)
            _, proposal = plan_batch(store, batch_id, ledger, rules, json.loads(skips.read_text(encoding="utf-8")))
            draft = proposal.vouchers[0]
            options = ExportOptions(PERIOD, "记", 1, used_numbers=frozenset({1}))
            with self.assertRaisesRegex(ExportError, "已使用"):
                export_workbook(template.read_bytes(), (draft,), ledger, options)
            with self.assertRaisesRegex(ExportError, "同一流水"):
                export_workbook(template.read_bytes(), (draft, draft), ledger, replace(options, used_numbers=frozenset()))

    def test_export_requires_review_and_clean_template_and_writes_literal_text(self):
        with TemporaryDirectory() as root:
            directory = Path(root)
            write_example(directory)
            store = StatementStore(directory / "statements.sqlite3")
            parsed = parse_ccb([ccb_row(remark="1+1")])
            batch = store.save(parsed)
            aux = (Auxiliary("项目", "029"), Auxiliary("自定义甲", "007"), Auxiliary("自定义乙", "008"))
            ledger = Ledger("book-ccb", 1, {"account-ccb": "1002001"}, {
                "1002001": Subject("1002001", "银行存款"), "1122001": Subject("1122001", "应收账款", ("项目", "自定义甲", "自定义乙")),
            }, frozenset(aux))
            rule = Rule("receive", 1, "book-ccb", "income", Assignment("1122001", aux), "={remark}", enabled=True)
            _, proposal = plan_batch(store, batch.id, ledger, (rule,))
            draft = proposal.vouchers[0]
            template = (directory / "synthetic-template.xlsx").read_bytes()
            options = ExportOptions(PERIOD, "记", 5, text_number=True)
            with self.assertRaisesRegex(ExportError, "核验"):
                export_workbook(template, (draft,), ledger, options, preview=False)
            options = replace(options, verified_template_hash=hashlib.sha256(template).hexdigest())
            with self.assertRaisesRegex(ExportError, "主数据"):
                export_workbook(template, (draft,), ledger, options, preview=False)
            ledger = replace(ledger, finance_verified=True)
            with self.assertRaisesRegex(ExportError, "已复核"):
                export_workbook(template, (draft,), ledger, options, preview=False)
            reviewed = review(draft, ledger)
            edit, issues = edit_entries(reviewed, (replace(reviewed.entries[0], amount_minor=999), *reviewed.entries[1:]), ledger)
            self.assertIsNone(edit.review_digest)
            self.assertEqual({issue.code for issue in issues}, {"bank_mismatch", "unbalanced"})
            stale = replace(reviewed, entries=(replace(reviewed.entries[0], summary="修改后"), *reviewed.entries[1:]))
            with self.assertRaisesRegex(ExportError, "内容未变化"):
                export_workbook(template, (stale,), ledger, options, preview=False)
            exported = export_workbook(template, (reviewed,), ledger, options, preview=False)
            workbook = load_workbook(BytesIO(exported.data), read_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["凭证模版"])
                sheet = workbook["凭证模版"]
                self.assertEqual((sheet["C3"].value, sheet["C3"].data_type), ("5", "s"))
                self.assertEqual((sheet["F3"].value, sheet["F3"].data_type), ("=1+1", "s"))
                self.assertEqual((sheet["N3"].value, sheet["N3"].data_type), ("029", "s"))
                self.assertEqual([sheet[f"{col}3"].value for col in "QRST"], ["自定义乙", "008", "自定义甲", "007"])
            finally:
                workbook.close()
            dirty = make_xlsx([EXPECTED_HEADERS, ("历史数据",)], title="凭证模版")
            with self.assertRaisesRegex(ExportError, "不是空白"):
                export_workbook(dirty, (reviewed,), ledger, options)

    def test_summarized_bank_entries_retain_each_transaction_reference(self):
        rows = parse_ccb([ccb_row(income=5), ccb_row(income=7, reference="000002")]).rows
        sources = tuple(Source(i, "ccb", "0306", PERIOD, 1, "流水", row.row_number, row.transaction)
                        for i, row in enumerate(rows, 1))
        ledger = Ledger("ccb", 1, {"0306": "1002001"}, {
            "1002001": Subject("1002001", "银行存款"), "1122001": Subject("1122001", "应收账款"),
        })
        rule = Rule("receive", 1, "ccb", "income", Assignment("1122001"), "收款", merge="summarize", enabled=True)
        proposal = propose(sources, ledger, (rule,))
        self.assertEqual(len(proposal.vouchers), 1)
        self.assertEqual([(e.side, e.amount_minor, e.transaction_ids) for e in proposal.vouchers[0].entries], [
            ("debit", 1200, (1, 2)), ("credit", 1200, (1, 2)),
        ])
