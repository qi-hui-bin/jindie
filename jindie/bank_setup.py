from __future__ import annotations

from dataclasses import asdict, dataclass, replace
from typing import Literal, Mapping

from .bank_profiles import BankProfile
from .excel_reader import StatementError
from .models import Issue, ParsedStatement, StatementRow, fingerprint

COLUMNS = {
    "own_account": "本方账号",
    "reference": "银行流水号",
    "transaction_date": "交易日期",
    "posting_date": "记账日期",
    "amount": "交易金额（正收负支）",
    "income": "收入金额",
    "expense": "支出金额",
    "currency": "币种",
    "counterparty_name": "对方户名",
    "counterparty_account": "对方账号",
    "summary": "摘要",
    "purpose": "用途",
    "remark": "备注",
    "balance": "余额",
}


def custom_profile(values: Mapping[str, str]) -> BankProfile:
    name = values["name"]
    if not name or len(name) > 80:
        raise StatementError("请输入流水格式名称，不超过 80 个字符。")
    mode: Literal["signed", "split"]
    if values["amount_mode"] == "signed":
        mode = "signed"
    elif values["amount_mode"] == "split":
        mode = "split"
    else:
        raise StatementError("请选择收支金额的表示方式。")
    used = set(COLUMNS) - ({"income", "expense"} if mode == "signed" else {"amount"})
    columns = {field: (values[field],) for field in COLUMNS if field in used and values[field]}
    if any(len(labels[0]) > 100 for labels in columns.values()):
        raise StatementError("Excel 表头名称不能超过 100 个字符。")
    required = {"own_account", "reference", "transaction_date"} | (
        {"amount"} if mode == "signed" else {"income", "expense"}
    )
    if missing := required - columns.keys():
        raise StatementError("请填写这些列的原表头：" + "、".join(COLUMNS[field] for field in COLUMNS if field in missing))
    if "currency" not in columns and values["cny_only"] != "yes":
        raise StatementError("请映射币种列，或明确确认此格式没有币种列且仅包含人民币。")
    sheet_name = values["sheet_name"]
    if len(sheet_name) > 31:
        raise StatementError("工作表名称不能超过 31 个字符。")
    row = values["header_row"]
    if row and (not row.isascii() or not row.isdecimal() or len(row) > 4 or not 1 <= int(row) <= 1000):
        raise StatementError("表头行须为 1 到 1000 之间的整数，或者留空自动定位。")
    profile = BankProfile(
        id="custom", version=1, name=name, columns=columns, amount_mode=mode,
        identify_by=tuple(columns), sheet_name=sheet_name or None, header_row=int(row) if row else None,
        currency_codes={"CNY": "CNY", "RMB": "CNY", "人民币": "CNY", "人民币元": "CNY"},
        default_currency=None if "currency" in columns else "CNY",
        allow_serial_dates=values["allow_serial_dates"] == "yes",
    )
    return replace(profile, id=f"custom-{fingerprint(asdict(profile))}")


@dataclass(frozen=True)
class MappingPreview:
    rows: tuple[StatementRow, ...]
    errors: tuple[tuple[int, Issue], ...]
    error_count: int
    income_count: int
    expense_count: int
    income_total: int
    expense_total: int

    @property
    def can_publish(self) -> bool:
        return self.error_count == 0 and self.income_count + self.expense_count > 0


def mapping_preview(statement: ParsedStatement) -> MappingPreview:
    rows: list[StatementRow] = []
    errors: list[tuple[int, Issue]] = []
    seen: set[str] = set()
    error_count = income_count = expense_count = income_total = expense_total = 0
    for row in statement.rows:
        issues = row.issues
        tx = row.transaction
        if tx is not None and not issues:
            if tx.bank_reference in seen:
                issues = (Issue("duplicate_reference", "样本内出现重复银行流水号。"),)
            elif tx.bank_reference is not None:
                seen.add(tx.bank_reference)
        if issues:
            error_count += 1
            if len(errors) < 12:
                errors.append((row.row_number, issues[0]))
        elif tx is not None:
            if len(rows) < 12:
                rows.append(row)
            if tx.direction == "income":
                income_count += 1
                income_total += tx.amount_minor
            else:
                expense_count += 1
                expense_total += tx.amount_minor
    return MappingPreview(tuple(rows), tuple(errors), error_count, income_count, expense_count, income_total, expense_total)
