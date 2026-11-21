from __future__ import annotations

import math
import struct
import warnings
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from xml.etree.ElementTree import ParseError
from xml.parsers.expat import ExpatError
from zipfile import BadZipFile, ZipFile

import xlrd
from defusedxml.common import DefusedXmlException
from openpyxl import load_workbook
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel
from xlrd.compdoc import CompDocError

from .models import RawValue

CellValue = RawValue | date | time | timedelta


class StatementError(ValueError):
    """A file or configuration error that prevents reliable parsing."""


@dataclass(frozen=True)
class ImportLimits:
    file_bytes: int = 20 * 1024 * 1024
    expanded_bytes: int = 128 * 1024 * 1024
    data_rows: int = 50_000
    columns: int = 256
    sheets: int = 32
    zip_entries: int = 1024

    def __post_init__(self) -> None:
        if min(self.file_bytes, self.expanded_bytes, self.data_rows, self.columns, self.sheets, self.zip_entries) < 1:
            raise ValueError("文件读取限制必须为正数。")


@dataclass(frozen=True)
class Cell:
    value: CellValue
    problem: str | None = None

    def raw(self) -> RawValue:
        if isinstance(self.value, (date, time)):
            return self.value.isoformat()
        if isinstance(self.value, timedelta):
            return str(self.value)
        if isinstance(self.value, float) and not math.isfinite(self.value):
            return str(self.value)
        return self.value


EMPTY = Cell(None)
ReadRows = Callable[[str], Iterator[tuple[Cell, ...]]]


@dataclass(frozen=True)
class WorksheetReader:
    sheet_names: tuple[str, ...]
    epoch: datetime
    legacy: bool
    rows: ReadRows


def _check_zip(data: bytes, limits: ImportLimits) -> None:
    with ZipFile(BytesIO(data)) as archive:
        parts = archive.infolist()
        names = set(archive.namelist())
        if len(parts) > limits.zip_entries or sum(part.file_size for part in parts) > limits.expanded_bytes:
            raise StatementError("工作簿解压后超过读取限制，请拆分后重试。")
        if any(part.flag_bits & 1 for part in parts):
            raise StatementError("文件已加密，请在本地另存为未加密的工作簿后重试。")
        if len(names) != len(parts) or not {"[Content_Types].xml", "xl/workbook.xml"} <= names:
            raise StatementError("该压缩文件不是有效的 Excel 工作簿。")


def _check_sheets(names: tuple[str, ...], limits: ImportLimits) -> None:
    if not names:
        raise StatementError("工作簿没有可读的数据工作表。")
    if len(names) > limits.sheets:
        raise StatementError("工作表数量超过读取限制，请仅保留需导入的工作表后重试。")


def _xlsx_rows(workbook, sheet_name: str, limits: ImportLimits) -> Iterator[tuple[Cell, ...]]:
    worksheet = workbook[sheet_name]
    worksheet.reset_dimensions()
    reader = worksheet.iter_rows()
    try:
        for number, row in enumerate(reader, 1):
            if len(row) > limits.columns:
                raise StatementError(f"工作表「{sheet_name}」第 {number} 行超过列数限制。")
            yield tuple(
                Cell(cell.value, "formula" if cell.data_type == "f" else "excel_error" if cell.data_type == "e" else None)
                for cell in row
            )
    finally:
        reader.close()


def _xls_rows(workbook, sheet_name: str, limits: ImportLimits) -> Iterator[tuple[Cell, ...]]:
    worksheet = workbook.sheet_by_name(sheet_name)
    if worksheet.ncols > limits.columns:
        raise StatementError(f"工作表「{sheet_name}」超过列数限制。")
    epoch = MAC_EPOCH if workbook.datemode else WINDOWS_EPOCH
    for number in range(worksheet.nrows):
        cells: list[Cell] = []
        for cell in worksheet.row(number):
            if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                cells.append(EMPTY)
            elif cell.ctype == xlrd.XL_CELL_ERROR:
                cells.append(Cell(xlrd.error_text_from_code.get(cell.value, "#ERROR"), "excel_error"))
            elif cell.ctype == xlrd.XL_CELL_DATE:
                if not math.isfinite(cell.value):
                    cells.append(Cell(str(cell.value), "excel_error"))
                    continue
                try:
                    cells.append(Cell(from_excel(cell.value, epoch)))
                except (ValueError, OverflowError):
                    cells.append(Cell(cell.value, "excel_error"))
            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                cells.append(Cell(bool(cell.value)))
            else:
                cells.append(Cell(cell.value))
        yield tuple(cells)


@contextmanager
def open_statement(filename: str, data: bytes, limits: ImportLimits) -> Iterator[WorksheetReader]:
    if not data or len(data) > limits.file_bytes:
        raise StatementError("文件为空或超过大小限制，请重新选择或拆分工作簿。")
    suffix = Path(filename).suffix.lower()
    ole = data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    zip_signature = data.startswith(b"PK\x03\x04")
    try:
        if suffix == ".xlsx" and zip_signature:
            _check_zip(data, limits)
            workbook = load_workbook(BytesIO(data), read_only=True, data_only=False, keep_links=False)
            try:
                # Chart sheets do not contain cells and are not candidates for matching.
                names = tuple(sheet.title for sheet in workbook.worksheets)
                _check_sheets(names, limits)
                # openpyxl converts out-of-range date cells to Excel errors; report those by row.
                with warnings.catch_warnings():
                    warnings.filterwarnings("ignore", message=r"Cell .* is marked as a date .*", category=UserWarning)
                    yield WorksheetReader(names, workbook.epoch, False, lambda name: _xlsx_rows(workbook, name, limits))
            finally:
                workbook.close()
        elif suffix == ".xls" and ole:
            workbook = xlrd.open_workbook(file_contents=data, on_demand=True, ragged_rows=True)
            try:
                names = tuple(workbook.sheet_names())
                _check_sheets(names, limits)
                epoch = MAC_EPOCH if workbook.datemode else WINDOWS_EPOCH
                yield WorksheetReader(names, epoch, True, lambda name: _xls_rows(workbook, name, limits))
            finally:
                workbook.release_resources()
        elif suffix == ".xlsx" and ole:
            raise StatementError("文件可能已加密，或实际为旧版 Excel。请在本地另存为未加密的 .xlsx 后重试。")
        else:
            raise StatementError("文件扩展名与内容不符或格式不支持，请提供有效的 .xls 或 .xlsx。")
    except StatementError:
        raise
    except (BadZipFile, xlrd.XLRDError, CompDocError, struct.error, ExpatError, ParseError, DefusedXmlException,
            KeyError, IndexError, ValueError, OverflowError, NotImplementedError) as exc:
        raise StatementError("未能读取工作簿，文件可能已损坏或加密。请在本地另存后重试。") from exc
