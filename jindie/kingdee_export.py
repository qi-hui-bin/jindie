from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from hashlib import sha256
from io import BytesIO
from typing import Iterable
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from defusedxml.common import DefusedXmlException
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils.exceptions import IllegalCharacterError, InvalidFileException

from .accounting import STANDARD_AUX, Ledger, Voucher, validate_voucher
from .excel_reader import ImportLimits, StatementError, _check_zip
from .models import Period

IMPORT_SHEET = "凭证模版"
PREVIEW_SHEET = "试算预览（不可导入）"
HEADERS = (
    "日期", "凭证字", "凭证号", "附件数", "分录序号", "摘要", "科目代码", "科目名称", "借方金额", "贷方金额",
    *STANDARD_AUX, "自定义辅助核算类别", "自定义辅助核算编码", "自定义辅助核算类别1", "自定义辅助核算编码1",
    "数量", "单价", "原币金额", "币别", "汇率",
)


class ExportError(ValueError):
    """The workbook or requested export is unsafe or invalid."""


@dataclass(frozen=True)
class ExportOptions:
    period: Period
    voucher_word: str
    first_number: int
    currency_code: str = "RMB"
    text_number: bool = False
    used_numbers: frozenset[int] = frozenset()
    verified_template_hash: str | None = None

    def __post_init__(self) -> None:
        if not self.voucher_word.strip() or not self.currency_code.strip() or (
            type(self.first_number) is not int or not 1 <= self.first_number <= 999_999_999_999_999
        ):
            raise ExportError("请设置凭证字、人民币币别代码及有效的起始凭证号。")
        if any(type(n) is not int or n < 1 for n in self.used_numbers):
            raise ExportError("既用凭证号必须为正整数。")


@dataclass(frozen=True)
class NumberedVoucher:
    number: int
    voucher: Voucher


@dataclass(frozen=True)
class WorkbookExport:
    data: bytes
    numbered: tuple[NumberedVoucher, ...]


ExcelValue = str | Decimal | int | None


def _number(vouchers: Iterable[Voucher], ledger: Ledger, options: ExportOptions, preview: bool) -> tuple[NumberedVoucher, ...]:
    ordered = list(vouchers)
    if not ordered:
        raise ExportError("没有可以导出的凭证。")
    seen: set[int] = set()
    for voucher in ordered:
        if issues := validate_voucher(voucher, ledger):
            raise ExportError("；".join(issue.message for issue in issues))
        if voucher.sources[0].period != options.period:
            raise ExportError("凭证期间和导出期间不一致。")
        if not preview and voucher.review_digest != voucher.digest(ledger):
            raise ExportError("实际导入版仅能导出已复核且内容未变化的凭证。")
        ids = {source.transaction_id for source in voucher.sources}
        if ids & seen:
            raise ExportError("同一流水不能关联到多张凭证。")
        seen.update(ids)
    numbered = tuple(
        NumberedVoucher(number, voucher)
        for number, voucher in enumerate(sorted(ordered, key=lambda item: item.order), options.first_number)
    )
    if numbered[-1].number > 999_999_999_999_999:
        raise ExportError("凭证编号超出 Excel 精度范围。")
    collisions = options.used_numbers.intersection(item.number for item in numbered)
    if collisions:
        raise ExportError(f"凭证号 {min(collisions)} 已使用，请调整起始号。")
    return numbered


def _rows(numbered: tuple[NumberedVoucher, ...], ledger: Ledger, options: ExportOptions) -> Iterable[tuple[ExcelValue, ...]]:
    for item in numbered:
        voucher = item.voucher
        for sequence, entry in enumerate(voucher.entries, 1):
            amount = Decimal(entry.amount_minor).scaleb(-2)
            auxiliary = {aux.category: aux.code for aux in entry.assignment.auxiliary}
            standard = tuple(auxiliary.pop(category, None) for category in STANDARD_AUX)
            custom = tuple(value for pair in sorted(auxiliary.items()) for value in pair)
            yield (
                voucher.sources[0].transaction.voucher_date.isoformat(), options.voucher_word,
                str(item.number) if options.text_number else item.number, 0, sequence,
                entry.summary, entry.assignment.subject, ledger.subjects[entry.assignment.subject].name,
                amount if entry.side == "debit" else None, amount if entry.side == "credit" else None,
                *standard, *custom, *((None,) * (4 - len(custom))), None, None,
                amount, options.currency_code, 1,
            )


def _write_cell(cell: Cell, value: ExcelValue) -> None:
    if isinstance(value, str):
        if len(value) > 32_767:
            raise ExportError("文本超过 Excel 单元格长度限制，请缩短摘要或检查主数据。")
        cell.value = value
        cell.data_type = "s"
    else:
        cell.value = value
    cell.number_format = "0.00" if isinstance(value, Decimal) else "@" if isinstance(value, str) else "General"


def _empty_template_sheet(workbook, limits: ImportLimits):
    if IMPORT_SHEET not in workbook.sheetnames or PREVIEW_SHEET in workbook.sheetnames:
        raise ExportError("请提供包含「凭证模版」数据页的干净模板。")
    sheet = workbook[IMPORT_SHEET]
    if sheet.max_column != len(HEADERS) or tuple(cell.value for cell in sheet[1]) != HEADERS:
        raise ExportError("「凭证模版」表头必须是顺序固定的 25 列。")
    if sheet.max_row > limits.data_rows + 1:
        raise ExportError("模板超出允许的行数，请改用没有多余空行的模板。")
    if any(cell.value is not None for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row) for cell in row):
        raise ExportError("模板数据页不是空白的，请勿使用含历史凭证的文件。")
    if sheet.merged_cells.ranges:
        raise ExportError("模板数据页包含合并单元格，无法写入分录。")
    return sheet


def validate_empty_template(template: bytes, limits: ImportLimits) -> None:
    if not template or len(template) > limits.file_bytes:
        raise ExportError("模板为空或超过大小限制。")
    try:
        _check_zip(template, limits)
        workbook = load_workbook(BytesIO(template), keep_links=False)
        try:
            _empty_template_sheet(workbook, limits)
        finally:
            workbook.close()
    except ExportError:
        raise
    except (StatementError, InvalidFileException, OSError, ValueError, KeyError, BadZipFile,
            DefusedXmlException, ParseError) as exc:
        raise ExportError("无法读取 Excel 模板，请用干净的目标账套模板重试。") from exc


def _verify(data: bytes, title: str, expected: list[tuple[ExcelValue, ...]]) -> None:
    workbook = load_workbook(BytesIO(data), read_only=True, keep_links=False)
    try:
        sheet = workbook[title]
        rows = sheet.iter_rows(min_col=1, max_col=len(HEADERS))
        try:
            if tuple(cell.value for cell in next(rows)) != HEADERS:
                raise ExportError("保存后读取的表头不匹配。")
            for actual, wanted in zip(rows, expected, strict=True):
                for cell, value in zip(actual, wanted, strict=True):
                    if isinstance(value, Decimal):
                        equal = cell.data_type == "n" and cell.value is not None and Decimal(str(cell.value)) == value
                    elif isinstance(value, str):
                        equal = cell.data_type == "s" and cell.value == value
                    else:
                        equal = cell.value == value and (value is None or type(cell.value) is int)
                    if not equal:
                        raise ExportError("保存后检查发现分录字段、编码或金额不匹配。")
        finally:
            rows.close()
    finally:
        workbook.close()


def export_workbook(
    template: bytes,
    vouchers: Iterable[Voucher],
    ledger: Ledger,
    options: ExportOptions,
    *,
    preview: bool = True,
) -> WorkbookExport:
    limits = ImportLimits()
    if not template or len(template) > limits.file_bytes:
        raise ExportError("模板为空或超过大小限制。")
    if not preview and sha256(template).hexdigest() != options.verified_template_hash:
        raise ExportError("实际导入版需要财务核验并登记的目标账套模板。")
    if not preview and not ledger.finance_verified:
        raise ExportError("实际导入版需要财务核验账套科目及辅助核算主数据。")
    numbered = _number(vouchers, ledger, options, preview)
    expected = list(_rows(numbered, ledger, options))
    if len(expected) >= 1_048_576:
        raise ExportError("分录数超过 Excel 行数上限，请分批处理。")
    try:
        _check_zip(template, limits)
        workbook = load_workbook(BytesIO(template), keep_links=False)
        try:
            sheet = _empty_template_sheet(workbook, limits)
            if preview:
                sheet.title = PREVIEW_SHEET
            else:
                for title in list(workbook.sheetnames):
                    if title != IMPORT_SHEET:
                        del workbook[title]
                sheet.sheet_state = "visible"
                workbook.active = sheet
            if sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row - 1)
            for row_index, values in enumerate(expected, 2):
                for position, value in enumerate(values, 1):
                    _write_cell(sheet.cell(row_index, position), value)
            output = BytesIO()
            workbook.save(output)
        finally:
            workbook.close()
        result = output.getvalue()
        _verify(result, PREVIEW_SHEET if preview else IMPORT_SHEET, expected)
    except ExportError:
        raise
    except (StatementError, InvalidFileException, IllegalCharacterError, OSError, ValueError, KeyError, StopIteration,
            BadZipFile, DefusedXmlException, ParseError) as exc:
        raise ExportError("无法读取或复查 Excel 模板，请用干净的目标账套模板重试。") from exc
    return WorkbookExport(result, numbered)


def reconciliation_workbook(
    exported: WorkbookExport, ledger: Ledger, options: ExportOptions,
    source_rows: Iterable[tuple[int, dict]], exclusions: list[dict],
) -> bytes:
    workbook = Workbook()
    overview = workbook.active
    overview.title = "说明"
    sources = workbook.create_sheet("流水底稿")
    details = workbook.create_sheet("凭证分录-11列")
    imported = workbook.create_sheet("精斗云导入-25列")
    summary = workbook.create_sheet("科目汇总")
    missing = workbook.create_sheet("待新增科目")
    doubts = workbook.create_sheet("存疑清单")

    def append(sheet, values):
        index = sheet.max_row + 1 if sheet["A1"].value is not None else 1
        for column, value in enumerate(values, 1):
            _write_cell(sheet.cell(index, column), value)

    append(overview, ("项目", "值"))
    append(overview, ("用途", "人工核对。精斗云请使用单独下载的单工作表文件。"))
    append(overview, ("账套", ledger.book_id))
    append(overview, ("期间", str(options.period)))
    append(overview, ("凭证张数", len(exported.numbered)))
    append(overview, ("凭证号段", f"{exported.numbered[0].number}–{exported.numbered[-1].number}"))
    append(sources, ("原始行", "日期", "对方户名", "摘要", "用途／附言", "收支", "金额", "流水号", "凭证号", "状态"))
    append(details, ("日期", "凭证字号", "摘要", "科目", "借方金额", "贷方金额", "附件", "原单据编号",
                     "制单人", "审核人", "制单日期"))
    append(imported, HEADERS)
    append(summary, ("科目代码", "科目名称", "借方合计", "贷方合计"))
    append(missing, ("原始行", "问题代码", "待核对"))
    append(doubts, ("原始行", "处理状态", "问题代码", "说明"))
    exceptions = {item["row_number"]: item for item in exclusions}
    source_numbers = {s.row_number: item.number for item in exported.numbered for s in item.voucher.sources}
    for row, tx in source_rows:
        number = source_numbers.get(row)
        excluded = exceptions.get(row, {})
        state = ("本次导出" if number else "已转重算" if excluded.get("recalculated_batch_id")
                 else "已跳过" if excluded.get("skipped") else "未纳入本次")
        amount = tx.get("amount_minor")
        append(sources, (row, tx.get("voucher_date"), tx.get("counterparty_name"), tx.get("summary"),
                         tx.get("purpose") or tx.get("message") or tx.get("remark"), tx.get("direction"),
                         Decimal(amount).scaleb(-2) if amount is not None else None,
                         tx.get("bank_reference"), number, state))
    for item in exclusions:
        state = "已转重算" if item.get("recalculated_batch_id") else "已跳过" if item.get("skipped") else "待核对"
        for issue in item["issues"]:
            append(doubts, (item["row_number"], state, issue["code"], issue["message"]))
            if issue["code"] in {"unknown_subject", "unknown_auxiliary", "missing_auxiliary"}:
                append(missing, (item["row_number"], issue["code"], issue["message"]))

    totals = defaultdict(lambda: [Decimal(0), Decimal(0)])
    for item in exported.numbered:
        references = {source.transaction_id: source.transaction.bank_reference for source in item.voucher.sources}
        for entry, row in zip(item.voucher.entries, _rows((item,), ledger, options), strict=True):
            append(imported, row)
            debit, credit = row[8:10]
            auxiliaries = " ".join(f"{aux.category}={aux.code}" for aux in entry.assignment.auxiliary)
            account = f"{row[6]} {row[7]}" + (f" [{auxiliaries}]" if auxiliaries else "")
            refs = "、".join(references[ref] for ref in entry.transaction_ids if references[ref])
            if len(refs) > 32_767:
                refs = f"{len(entry.transaction_ids)} 笔，见流水底稿"
            append(details, (row[0], f"{row[1]}-{row[2]}", row[5], account, debit, credit, 0, refs, None, None, None))
            totals[row[6]][0] += debit or 0
            totals[row[6]][1] += credit or 0
    for code, (debit, credit) in sorted(totals.items()):
        append(summary, (code, ledger.subjects[code].name, debit, credit))
    append(overview, ("借方合计", sum(debit for debit, _ in totals.values())))
    append(overview, ("贷方合计", sum(credit for _, credit in totals.values())))
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for column in sheet.columns:
            width = max((len(str(cell.value)) for cell in column if cell.value is not None), default=6)
            sheet.column_dimensions[column[0].column_letter].width = min(max(12, width * 1.6 + 2), 64)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()
