import unittest
from dataclasses import replace
from datetime import date, datetime, time
from io import BytesIO
from zipfile import ZipFile

from openpyxl.utils.datetime import MAC_EPOCH, to_excel

from jindie.bank_import import parse_statement
from jindie.bank_profiles import CCB_0306, BankProfile
from jindie.excel_reader import ImportLimits, StatementError
from jindie.models import BookAccount, Period

from .fixtures import BOC_ACCOUNT, CCB_ACCOUNT, CCB_HEADERS, boc_rows, ccb_row, ccb_rows, make_xls, make_xlsx

PERIOD = Period(2026, 11)
BOC = BookAccount("book-boc", "account-boc", BOC_ACCOUNT)
CCB = BookAccount("book-ccb", "account-ccb", CCB_ACCOUNT)


def parse_ccb(rows, **settings):
    account = settings.pop("account", CCB)
    return parse_statement("synthetic.xlsx", make_xlsx([CCB_HEADERS, *rows]), account=account, period=PERIOD, **settings)


class BankImportTests(unittest.TestCase):
    def test_two_bank_profiles_on_synthetic_xls_and_xlsx(self):
        boc = parse_statement("test.xls", make_xls(boc_rows()), account=BOC, period=PERIOD)
        self.assertEqual((boc.sheet, boc.header_row, len(boc.rows)), ("Sheet1", 9, 49))
        self.assertEqual(boc.warnings[0].code, "xls_cached_values")
        for i, row in enumerate(boc.rows):
            with self.subTest(bank="BOC", row=row.row_number):
                self.assertFalse(row.issues)
                tx = row.transaction
                self.assertIsNotNone(tx)
                self.assertEqual(row.row_number, 10 + i)
                self.assertEqual(tx.bank_reference, f"{i + 1:06d}")
                self.assertEqual(tx.direction, "income" if i < 37 else "expense")
                self.assertEqual(tx.amount_minor, 150_000 + i * 100 if i < 37 else (900, 90, 330, 250, 200, 400)[(i - 37) % 6])
                self.assertEqual(tx.transaction_date, date(2026, 11, min(12 + i // 7, 18)))
                self.assertEqual(tx.counterparty_name, f"合成对方{i}")
        for extension, generate in (("xls", make_xls), ("xlsx", make_xlsx)):
            ccb = parse_statement(f"test.{extension}", generate(ccb_rows(), "合成建行0306"), account=CCB, period=PERIOD)
            self.assertEqual((ccb.sheet, ccb.header_row, len(ccb.rows)), ("合成建行0306", 1, 91))
            for i, row in enumerate(ccb.rows):
                with self.subTest(bank="CCB", extension=extension, row=row.row_number):
                    self.assertFalse(row.issues)
                    tx = row.transaction
                    self.assertIsNotNone(tx)
                    self.assertEqual(tx.direction, "income" if i < 32 else "expense")
                    self.assertEqual(tx.amount_minor, (100 + i) * 100)
                    self.assertEqual(tx.bank_reference, f"{i + 1:06d}")
                    self.assertEqual(tx.transaction_date, date(2026, 11, 2 + i * 18 // 90))
            self.assertIsNone(ccb.rows[0].raw["expense"])
            self.assertEqual(ccb.rows[-1].raw["income"], 0)

    def test_errors_are_kept_with_excel_row_numbers(self):
        rows = [
            ccb_row(income=0, expense=0),
            ccb_row(account="0000000000001111"),
            ccb_row(income="1,234.567"),
            ccb_row(income="1.00000000000000000000000000000000000001"),
            ccb_row(income="=8+2"),
            ccb_row(tx="2026-11-31"),
            (),
            ccb_row(tx="2026-12-01", posting="2026-12-01"),
            ("合计", 0, 100),
            ccb_row(account=6222123456789091),
            ccb_row(income="1,234.50", reference=None, remark=" 合成备注 "),
        ]
        parsed = parse_ccb(rows)
        self.assertEqual([(row.row_number, row.issues[0].code) for row in parsed.rows], [
            (2, "direction_conflict"), (3, "account_mismatch"), (4, "amount_precision"),
            (5, "amount_precision"), (6, "formula"), (7, "invalid_date"), (9, "out_of_period"),
            (11, "unsafe_identifier"), (12, "missing_reference"),
        ])
        self.assertEqual(parsed.rows[-1].transaction.amount_minor, 123450)
        self.assertEqual(parsed.rows[-1].transaction.remark, " 合成备注 ")

    def test_posting_date_and_account_confirmations_are_explicit(self):
        rows = [ccb_row(tx="2026-10-31T9:05:00")]
        self.assertEqual(parse_ccb(rows).rows[0].issues[0].code, "date_conflict")
        self.assertEqual(parse_ccb(rows, posting_preference="transaction").rows[0].issues[0].code, "out_of_period")
        tx = parse_ccb(rows, posting_preference="posting").rows[0].transaction
        self.assertEqual((tx.voucher_date, tx.transaction_time), (date(2026, 11, 2), time(9, 5)))
        self.assertEqual(parse_ccb([ccb_row(account="****0306")]).rows[0].issues[0].code, "account_confirmation")
        self.assertEqual(parse_ccb([ccb_row(account="****0306")], confirmed_binding=True).rows[0].transaction.own_account, CCB_ACCOUNT)
        self.assertEqual(parse_ccb([ccb_row(account="****9130")], confirmed_binding=True).rows[0].issues[0].code, "account_mismatch")
        rows = boc_rows()
        bad_payment = list(rows[9])
        bad_payment[6], bad_payment[8] = bad_payment[8], bad_payment[6]
        rows[9] = tuple(bad_payment)
        self.assertEqual(parse_statement("boc.xls", make_xls(rows), account=BOC, period=PERIOD).rows[0].issues[0].code, "direction_conflict")

    def test_ccb_accepts_compact_date_with_clock_in_legacy_xls(self):
        rows = [
            CCB_HEADERS,
            ccb_row(tx="20261104 08:57:30", posting="20261104"),
            ccb_row(tx="20261131 08:57:30", posting="20261104", reference="000002"),
            ccb_row(tx="20261104 24:00:00", posting="20261104", reference="000003"),
        ]
        statement = parse_statement("ccb.xls", make_xls(rows), account=CCB, period=PERIOD)
        self.assertFalse(statement.rows[0].issues)
        tx = statement.rows[0].transaction
        self.assertEqual((tx.transaction_date, tx.transaction_time, tx.voucher_date),
                         (date(2026, 11, 4), time(8, 57, 30), date(2026, 11, 4)))
        self.assertEqual([row.issues[0].code for row in statement.rows[1:]], ["invalid_date", "invalid_date"])

    def test_reordered_custom_mapping_uses_workbook_epoch(self):
        rows = [
            ("账号", "本币金额\nAmount", "业务日期(Date)", "原流水", "附言"),
            (CCB_ACCOUNT, "-44.20", to_excel(datetime(2026, 11, 12, 12, 0), MAC_EPOCH), "00029", "自定义"),
        ]
        profile = BankProfile(
            id="custom", version=1, name="自定义", amount_mode="signed", currency_codes={},
            columns={"own_account": ("账号",), "reference": ("原流水",), "amount": ("本币金额",),
                     "transaction_date": ("业务日期",), "remark": ("附言",)},
            default_currency="CNY", allow_serial_dates=True, ignore_english_suffix=True,
        )
        parsed = parse_statement("custom.xlsx", make_xlsx(rows, epoch=MAC_EPOCH), account=CCB, period=PERIOD, profiles=(profile,))
        tx = parsed.rows[0].transaction
        self.assertIsNotNone(tx)
        self.assertEqual((tx.transaction_date, tx.transaction_time), (date(2026, 11, 12), time(12)))
        self.assertEqual((tx.bank_reference, tx.amount_minor, tx.direction), ("00029", 4420, "expense"))

    def test_requires_choice_if_more_than_one_sheet_or_profile_matches(self):
        data = make_xlsx([CCB_HEADERS, ccb_row()], "第一张", second_sheet=True)
        with self.assertRaisesRegex(StatementError, "多处匹配"):
            parse_statement("test.xlsx", data, account=CCB, period=PERIOD)
        result = parse_statement("test.xlsx", data, account=CCB, period=PERIOD, sheet="第一张")
        self.assertEqual(len(result.rows), 1)
        with self.assertRaisesRegex(StatementError, "多处匹配"):
            parse_statement("test.xlsx", data, account=CCB, period=PERIOD, sheet="第一张",
                            profiles=(CCB_0306, replace(CCB_0306, id="second")))

    def test_limits_and_invalid_workbooks_have_actionable_errors(self):
        data = make_xlsx([CCB_HEADERS, ccb_row(), ccb_row(reference="000002")])
        for name, payload, limits, message in (
            ("wrong.xls", data, ImportLimits(), "扩展名"),
            ("limited.xlsx", data, ImportLimits(data_rows=1), "数据行数"),
            ("limited.xlsx", data, ImportLimits(expanded_bytes=100), "解压"),
            ("limited.xlsx", data, ImportLimits(file_bytes=10), "大小"),
            ("corrupt.xlsx", data[:-25], ImportLimits(), "损坏"),
        ):
            with self.subTest(name=name, message=message), self.assertRaisesRegex(StatementError, message):
                parse_statement(name, payload, account=CCB, period=PERIOD, limits=limits)
        unrelated = BytesIO()
        with ZipFile(unrelated, "w") as zip_file:
            zip_file.writestr("not-excel.txt", "example")
        with self.assertRaisesRegex(StatementError, "不是有效"):
            parse_statement("wrong.xlsx", unrelated.getvalue(), account=CCB, period=PERIOD)
