from io import BytesIO
from unittest import TestCase

from openpyxl import load_workbook

from jindie.accounting import Assignment, FixedEntry, Ledger, Rule, Source, Subject, propose
from jindie.bank_import import parse_statement
from jindie.kingdee_export import ExportOptions, HEADERS, export_workbook, reconciliation_workbook
from jindie.models import BookAccount, Period

from .fixtures import make_xlsx


class SopTests(TestCase):
    def test_cib_payroll_and_both_workbooks(self):
        account = BookAccount("test", "cib", "000001")
        period = Period(2026, 11)
        headers = ("账号", "记账日期", "交易时间", "借方金额(支出)", "贷方金额(收入)", "币种", "唯一流水编号")
        raw = make_xlsx([
            ("账户明细",), headers,
            ("000001", "2026-11-13", "2026-11-13 17:01:54", 30, 0, "人民币", "001"),
            ("000001", "2026-11-13", "2026-11-13 17:01:55", 70, 0, "人民币", "002"),
        ])
        parsed = parse_statement("cib.xlsx", raw, account=account, period=period)
        self.assertEqual(parsed.profile_id, "cib")
        self.assertTrue(all(not row.issues and row.transaction for row in parsed.rows))
        sources = tuple(Source(i, "test", "cib", period, 1, parsed.sheet, row.row_number, row.transaction)
                        for i, row in enumerate(parsed.rows, 1))
        ledger = Ledger("test", 1, {"cib": "100201"}, {
            "100201": Subject("100201", "银行"),
            "221101": Subject("221101", "工资"),
            "222112": Subject("222112", "个税"),
        })
        rule = Rule(
            "工资", 1, "test", "expense", None, "发放工资", account_id="cib", enabled=True, merge="bank_detail",
            fixed_entries=(FixedEntry(Assignment("221101"), "debit", 12_000),
                           FixedEntry(Assignment("222112"), "credit", 2_000)),
            bank_references=("001", "002"),
        )
        result = propose(sources, ledger, (rule,))
        self.assertEqual(len(result.vouchers), 1)
        self.assertEqual([entry.amount_minor for entry in result.vouchers[0].entries], [12000, 2000, 3000, 7000])
        self.assertEqual(propose(sources[:1], ledger, (rule,)).exclusions[0].issues[0].code, "incomplete_fixed_group")
        options = ExportOptions(period, "记", 249)
        exported = export_workbook(make_xlsx([HEADERS], title="凭证模版"), result.vouchers, ledger, options)
        report = reconciliation_workbook(exported, ledger, options,
                                         ((s.row_number, s.transaction.payload()) for s in sources), [])
        w = load_workbook(BytesIO(report), read_only=True)
        try:
            self.assertEqual(w.sheetnames, ["说明", "流水底稿", "凭证分录-11列", "精斗云导入-25列",
                                            "科目汇总", "待新增科目", "存疑清单"])
            self.assertEqual(list(w["精斗云导入-25列"].values)[0], HEADERS)
            self.assertEqual(w["流水底稿"]["H2"].value, "001")
        finally:
            w.close()
