from __future__ import annotations

import hashlib
import math
import re
from contextlib import closing
from dataclasses import asdict, dataclass
from datetime import date, datetime, time
from decimal import Decimal
from itertools import islice
from typing import Literal

from openpyxl.utils.datetime import from_excel

from .bank_profiles import BUILTIN_PROFILES, BankProfile, normalize_header
from .excel_reader import EMPTY, Cell, ImportLimits, StatementError, WorksheetReader, open_statement
from .models import BookAccount, Direction, Issue, ParsedStatement, Period, StatementRow, StatementTransaction, canonical_json

PostingPreference = Literal["transaction", "posting"]
AMOUNT_TEXT = re.compile(r"[+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?")
DATE_TEXT = re.compile(
    r"(?P<date>[0-9]{8}|[0-9]{4}[-/.][0-9]{1,2}[-/.][0-9]{1,2})"
    r"(?:[ T]+(?P<time>[0-9]{1,2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,6})?)?))?"
)
MASK = re.compile(r"[0-9*＊xX•-]+")
MINOR_LIMIT = 2**63 - 1


class RowProblem(ValueError):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.issue = Issue(code, message)


@dataclass(frozen=True)
class HeaderMatch:
    profile: BankProfile
    sheet: str
    row: int
    columns: dict[str, int]
    labels: dict[str, str]
    ambiguous: bool


def _text(cell: Cell) -> str | None:
    if cell.value is None or isinstance(cell.value, str) and not cell.value.strip():
        return None
    return str(cell.raw())


def _identifier(cell: Cell, label: str) -> str | None:
    value = cell.value
    if isinstance(value, (int, float)):
        if isinstance(value, bool) or not math.isfinite(value) or int(value) != value or not 0 <= value < 10**15:
            raise RowProblem("unsafe_identifier", f"{label}是无法可靠恢复的数值，请以文本格式重新获取原值。")
        return str(int(value))
    if value is not None and not isinstance(value, str):
        raise RowProblem("unsafe_identifier", f"{label}不是文本，请核对原文件。")
    text = _text(cell)
    if text and re.fullmatch(r"[0-9]+(?:\.[0-9]+)?[Ee][+-]?[0-9]+", text.strip()):
        raise RowProblem("unsafe_identifier", f"{label}使用科学计数法，请核对原文件并提供完整文本。")
    return text.strip() if text else None


def _minor(cell: Cell, label: str) -> int | None:
    value = cell.value
    if _text(cell) is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        if len(value) > 80 or not AMOUNT_TEXT.fullmatch(value):
            raise RowProblem("invalid_amount", f"{label}不是有效金额，请核对该单元格。")
        number = Decimal(value.replace(",", ""))
    elif isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
        number = Decimal(str(value))
    else:
        raise RowProblem("invalid_amount", f"{label}不是有效金额，请核对该单元格。")
    if not -MINOR_LIMIT <= number * 100 <= MINOR_LIMIT:
        raise RowProblem("invalid_amount", f"{label}超出可保存范围。")
    # Use the original decimal digits rather than rounding before checking precision.
    parts = number.as_tuple()
    if isinstance(parts.exponent, int) and parts.exponent < -2 and any(parts.digits[parts.exponent + 2:]):
        raise RowProblem("amount_precision", f"{label}有超过两位有效小数，请人工复核。")
    return int(number * 100)


def _clock_text(text: str) -> time:
    if re.match(r"^[0-9]:", text):
        text = "0" + text
    if not re.fullmatch(r"[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,6})?)?", text):
        raise ValueError("invalid time")
    return time.fromisoformat(text)


def _date_value(cell: Cell, label: str, profile: BankProfile, epoch: datetime) -> tuple[date, time | None]:
    value = cell.value
    if isinstance(value, datetime):
        return value.date(), value.time()
    if isinstance(value, date):
        return value, None
    try:
        if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
            if int(value) == value and 10000101 <= value <= 99991231:
                text = str(int(value))
            elif profile.allow_serial_dates:
                converted = from_excel(value, epoch)
                if isinstance(converted, datetime):
                    return converted.date(), converted.time() if value % 1 else None
                raise ValueError("not a date")
            else:
                raise ValueError("serial dates are disabled")
        else:
            text = (value or "").strip() if isinstance(value, str) else ""
        if match := DATE_TEXT.fullmatch(text):
            token = match["date"]
            if token.isdigit():
                source_date = date(int(token[:4]), int(token[4:6]), int(token[6:]))
            else:
                year, month, day = re.split(r"[-/.]", token)
                source_date = date(int(year), int(month), int(day))
            return source_date, _clock_text(match["time"]) if match["time"] else None
    except (ValueError, OverflowError):
        raise RowProblem("invalid_date", f"{label}无效，请核对交易或记账日期。") from None
    raise RowProblem("invalid_date", f"{label}缺失或格式不受支持，请核对原文件。")


def _time_value(cell: Cell, profile: BankProfile, epoch: datetime) -> time | None:
    if _text(cell) is None:
        return None
    value = cell.value
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    if profile.allow_serial_dates and isinstance(value, (int, float)) and not isinstance(value, bool) and 0 <= value < 1:
        converted = from_excel(value, epoch)
        if isinstance(converted, time):
            return converted
    if isinstance(value, str):
        try:
            text = value.strip()
            match = DATE_TEXT.fullmatch(text)
            if match and match["time"]:
                return _date_value(cell, "交易时间", profile, epoch)[1]
            return _clock_text(text)
        except ValueError:
            pass
    raise RowProblem("invalid_time", "交易时间无效，请核对原文件。")


def _match_headers(reader: WorksheetReader, profiles: tuple[BankProfile, ...], sheet: str | None) -> list[HeaderMatch]:
    if sheet is not None and sheet not in reader.sheet_names:
        raise StatementError("所选工作表不存在，请重新选择。")
    matches: list[HeaderMatch] = []
    for name in (sheet,) if sheet else reader.sheet_names:
        active = tuple(p for p in profiles if p.sheet_name in (None, name))
        if not active:
            continue
        stop = max(p.header_row or p.header_scan_rows for p in active)
        with closing(reader.rows(name)) as rows:
            for number, row in enumerate(islice(rows, stop), 1):
                for profile in active:
                    if (profile.header_row is not None and number != profile.header_row) or (
                        profile.header_row is None and number > profile.header_scan_rows
                    ):
                        continue
                    aliases = {
                        normalize_header(alias, profile.ignore_english_suffix): field
                        for field, values in profile.columns.items() for alias in values
                    }
                    found: dict[str, int] = {}
                    labels: dict[str, str] = {}
                    ambiguous = False
                    for index, cell in enumerate(row):
                        if cell.problem or not isinstance(cell.value, str):
                            continue
                        mapped = aliases.get(normalize_header(cell.value, profile.ignore_english_suffix))
                        if mapped is not None:
                            ambiguous |= mapped in found
                            found[mapped], labels[mapped] = index, cell.value
                    if profile.required_fields <= found.keys():
                        matches.append(HeaderMatch(profile, name, number, found, labels, ambiguous))
    return matches


def _header_account(row: tuple[Cell, ...], label: str | None, ignore_english_suffix: bool = False) -> Cell | None:
    if not label:
        return None
    for index, cell in enumerate(row):
        if not cell.problem and isinstance(cell.value, str):
            key, separator, rest = re.sub("：", ":", cell.value).partition(":")
            if normalize_header(key, ignore_english_suffix) == normalize_header(label, ignore_english_suffix):
                return Cell(rest) if separator and rest.strip() else next(
                    (candidate for candidate in row[index + 1:] if _text(candidate) is not None or candidate.problem),
                    EMPTY,
                )
    return None


def _full_account(cell: Cell, expected: str, confirmed_binding: bool) -> str:
    number = _identifier(cell, "本方账号")
    if number and re.fullmatch(r"[0-9]+", number):
        if number != expected:
            raise RowProblem("account_mismatch", "本方账号与所选账户不一致，请选择正确的账套和银行账户。")
        return number
    if number:
        if not MASK.fullmatch(number) or not re.search(r"[*＊xX•]", number):
            raise RowProblem("invalid_account", "本方账号格式无效，请核对原文件。")
        prefix = re.match(r"[0-9]*", number)
        suffix = re.search(r"[0-9]*$", number)
        if (prefix and not expected.startswith(prefix.group())) or (suffix and not expected.endswith(suffix.group())):
            raise RowProblem("account_mismatch", "本方账号可见部分与所选账户不一致，请核对原文件。")
    if not confirmed_binding:
        raise RowProblem("account_confirmation", "本方账号缺失或被遮蔽，请核对后确认绑定到所选账户。")
    return expected


def _signed_amount(profile: BankProfile, cells: dict[str, Cell]) -> tuple[Direction, int]:
    if profile.amount_mode == "split":
        income = _minor(cells.get("income", EMPTY), "收入") or 0
        expense = _minor(cells.get("expense", EMPTY), "支出") or 0
        if min(income, expense) < 0 or (income > 0) == (expense > 0):
            raise RowProblem("direction_conflict", "收支两列须仅一列为正，另一列为零或空；请人工核对。")
        return ("income", income) if income else ("expense", expense)
    amount = _minor(cells.get("amount", EMPTY), "交易金额")
    if not amount:
        raise RowProblem("invalid_amount", "交易金额缺失或为零，请核对原文件。")
    direction: Direction = "income" if amount > 0 else "expense"
    if profile.amount_mode == "coded_signed":
        code = (_text(cells.get("direction_code", EMPTY)) or "").strip()
        configured = profile.direction_codes.get(code)
        if configured is None or configured != direction:
            raise RowProblem("direction_conflict", "交易类型未知或与金额正负不一致，请人工核对。")
    return direction, abs(amount)


def _parse_transaction(
    cells: dict[str, Cell],
    profile: BankProfile,
    epoch: datetime,
    account: BookAccount,
    period: Period,
    posting_preference: PostingPreference | None,
    confirmed_binding: bool,
) -> StatementTransaction:
    def cell(field: str) -> Cell:
        return cells.get(field, EMPTY)

    direction, amount = _signed_amount(profile, cells)
    own_account = _full_account(cell("own_account"), account.number, confirmed_binding)
    transaction_date, combined_time = _date_value(cell("transaction_date"), "交易日期", profile, epoch)
    transaction_time = _time_value(cell("transaction_time"), profile, epoch) or combined_time
    posting_date = _date_value(cell("posting_date"), "记账日期", profile, epoch)[0] if "posting_date" in profile.columns else None
    if posting_date and posting_date != transaction_date and posting_preference is None:
        raise RowProblem("date_conflict", "交易日期与记账日期不同，请确认该账套使用哪个日期。")
    voucher_date = posting_date if posting_date and posting_preference == "posting" else transaction_date
    if not period.contains(voucher_date):
        raise RowProblem("out_of_period", "候选凭证日期不在所选会计期间，请检查期间或日期配置。")
    payer = _identifier(cell("payer_account"), "付款人账号")
    payee = _identifier(cell("payee_account"), "收款人账号")
    side, other = (payee, payer) if direction == "income" else (payer, payee)
    if other == own_account or side and side.isascii() and side.isdigit() and side != own_account:
        raise RowProblem("direction_conflict", "付款人与收款人账号和本方收支方向不一致，请人工核对。")
    name_field = "payer_name" if direction == "income" else "payee_name"
    raw_currency = _text(cell("currency"))
    currency = profile.currency_codes.get(raw_currency.strip()) if raw_currency else profile.default_currency
    if not currency:
        raise RowProblem("unknown_currency", "币种缺失或未配置映射，请检查银行格式。")
    return StatementTransaction(
        bank_reference=_identifier(cell("reference"), "银行流水号"),
        own_account=own_account,
        counterparty_account=other or _identifier(cell("counterparty_account"), "对方账号"),
        counterparty_name=_text(cell(name_field)) or _text(cell("counterparty_name")),
        transaction_date=transaction_date,
        transaction_time=transaction_time,
        posting_date=posting_date,
        voucher_date=voucher_date,
        direction=direction,
        amount_minor=amount,
        currency=currency,
        balance_minor=_minor(cell("balance"), "余额"),
        summary=_text(cell("summary")),
        purpose=_text(cell("purpose")),
        message=_text(cell("message")),
        remark=_text(cell("remark")),
    )


def parse_statement(
    filename: str,
    data: bytes,
    *,
    account: BookAccount,
    period: Period,
    profiles: tuple[BankProfile, ...] = BUILTIN_PROFILES,
    sheet: str | None = None,
    posting_preference: PostingPreference | None = None,
    confirmed_binding: bool = False,
    limits: ImportLimits = ImportLimits(),
) -> ParsedStatement:
    if posting_preference not in (None, "transaction", "posting"):
        raise StatementError("记账日期配置无效。")
    with open_statement(filename, data, limits) as reader:
        matches = _match_headers(reader, profiles, sheet)
        if not matches:
            raise StatementError("未识别到银行表头，请检查工作表或配置字段映射后重试。")
        if len(matches) > 1:
            choices = "；".join(f"{m.profile.name} / {m.sheet} / 第 {m.row} 行" for m in matches[:8])
            raise StatementError(f"有多处匹配的银行表头，请选择工作表、配置或固定表头行：{choices}")
        match = matches[0]
        if match.ambiguous:
            raise StatementError("同一个字段对应表头中的多列，请调整字段映射。")
        profile = match.profile
        ignored = {normalize_header(label, profile.ignore_english_suffix) for label in profile.ignored_row_labels}
        records: list[StatementRow] = []
        warnings = (Issue("xls_cached_values", ".xls 仅读取已保存的结果，无法保证识别每个公式单元格；请核对或提供值版。"),) if reader.legacy else ()
        header_account: Cell | None = None
        with closing(reader.rows(match.sheet)) as rows:
            for number, row in enumerate(rows, 1):
                if number < match.row:
                    if profile.account_label_row in (None, number):
                        found = _header_account(row, profile.account_label, profile.ignore_english_suffix)
                        if found is not None:
                            if header_account is not None and header_account != found:
                                raise StatementError("发现多个不同的查询账号，请核对银行文件。")
                            header_account = found
                    continue
                if number == match.row:
                    continue
                if number - match.row > limits.data_rows:
                    raise StatementError("数据行数超过读取限制，请拆分工作簿后重试。")
                first = next((c for c in row if _text(c) is not None or c.problem), None)
                if first is None:
                    continue
                cells = {k: row[i] if i < len(row) else EMPTY for k, i in match.columns.items()}
                if not first.problem and normalize_header(_text(first) or "", profile.ignore_english_suffix) in ignored and (
                    _text(cells.get("reference", EMPTY)) is None and
                    cells.get("transaction_date", EMPTY) in (EMPTY, first) and not any(c.problem for c in cells.values())
                ):
                    continue
                cells.setdefault("own_account", header_account or EMPTY)
                raw = {key: value.raw() for key, value in cells.items()}
                problems = tuple(
                    Issue(c.problem, f"「{match.labels.get(field, profile.account_label or field)}」含公式或 Excel 错误，请另存为普通值后重试。")
                    for field, c in cells.items() if c.problem is not None
                )
                tx = None
                if not problems:
                    try:
                        tx = _parse_transaction(cells, profile, reader.epoch, account, period, posting_preference, confirmed_binding)
                    except RowProblem as exc:
                        problems = (exc.issue,)
                    else:
                        if tx.bank_reference is None:
                            problems = (Issue("missing_reference", "缺少银行流水号，须人工查重后才能继续。"),)
                records.append(StatementRow(number, raw, problems, tx))
    return ParsedStatement(
        filename=re.split(r"[/\\]", filename)[-1],
        file_hash=hashlib.sha256(data).hexdigest(),
        sheet=match.sheet,
        header_row=match.row,
        profile_id=profile.id,
        profile_version=profile.version,
        profile_json=canonical_json(asdict(profile)),
        account=account,
        period=period,
        posting_preference=posting_preference,
        confirmed_binding=confirmed_binding,
        warnings=warnings,
        rows=tuple(records),
    )
