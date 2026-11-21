from __future__ import annotations

import re
from collections.abc import Iterator, Mapping
from hashlib import sha256
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook

from .accounting import STANDARD_AUX
from .bank_profiles import BOC_9130, CCB_0306, CIB
from .excel_reader import ImportLimits, StatementError, WorksheetReader, open_statement
from .kingdee_export import validate_empty_template

SETUP_LIMITS = ImportLimits(file_bytes=10 * 1024 * 1024)
FORMATS = {"中行": BOC_9130.id, "建行": CCB_0306.id, "兴业": CIB.id}
DATE_PREFERENCES = {"交易日": "transaction", "记账日": "posting"}
SUBJECTS = "科目"
AUXILIARY = "辅助核算"
ACCOUNTS = "银行账户"
HEADERS = {
    SUBJECTS: ("科目编码", "科目名称", "必填辅助类别", "数量核算"),
    AUXILIARY: ("辅助类别", "辅助编码"),
    ACCOUNTS: ("账户标识", "账户名称", "完整账号", "银行科目编码", "流水格式", "日期口径"),
}


def setup_workbook(config: dict | None = None) -> bytes:
    workbook = Workbook()
    workbook.remove(workbook.active)
    content: dict[str, list[tuple[str, ...]]] = {}
    if config is not None:
        ledger = config["ledger"]
        format_names = {value: key for key, value in FORMATS.items()}
        date_names = {value: key for key, value in DATE_PREFERENCES.items()}
        content = {
            SUBJECTS: [
                (row["code"], row["name"], "、".join(row["required_aux"]), "是" if row["quantity_required"] else "否")
                for row in ledger["subjects"]
            ],
            AUXILIARY: [(row["category"], row["code"]) for row in ledger["auxiliary"]],
            ACCOUNTS: [
                (item["id"], item["label"], item["number"], ledger["bank_subjects"][item["id"]],
                 "自定义" if "bank_profile" in item else format_names[item["profile_id"]] if item["profile_id"] else "自动",
                 date_names[item["posting_preference"]] if item["posting_preference"] else "")
                for item in config["web"]["accounts"]
            ],
        }
    for title, headers in HEADERS.items():
        sheet = workbook.create_sheet(title)
        sheet.append(headers)
        for cell in sheet[1]:
            sheet.column_dimensions[cell.column_letter].number_format = "@"
        for values in content.get(title, []):
            sheet.append(values)
            for cell in sheet[sheet.max_row]:
                cell.data_type = "s"
    with BytesIO() as output:
        workbook.save(output)
        return output.getvalue()


def _rows(reader: WorksheetReader, title: str) -> Iterator[tuple[int, tuple[str, ...]]]:
    if title not in reader.sheet_names:
        raise StatementError(f"主数据工作簿缺少「{title}」表。")
    headers = HEADERS[title]
    rows = reader.rows(title)
    try:
        first = next(rows, ())
        if tuple(cell.value for cell in first[:len(headers)]) != headers or any(
            cell.value is not None for cell in first[len(headers):]
        ):
            raise StatementError(f"「{title}」表的第 1 行表头不符，请使用下载的表格。")
        for number, cells in enumerate(rows, 2):
            if number > SETUP_LIMITS.data_rows + 1:
                raise StatementError(f"「{title}」表超过行数限制。")
            if any(cell.value is not None or cell.problem for cell in cells[len(headers):]):
                raise StatementError(f"「{title}」第 {number} 行有表头以外的数据。")
            if not any(cell.value is not None or cell.problem for cell in cells):
                continue
            values = []
            for header, cell in zip(headers, cells):
                if cell.problem or (cell.value is not None and not isinstance(cell.value, str)):
                    raise StatementError(f"「{title}」第 {number} 行的「{header}」须为文本，不能是数字、公式或错误值。")
                values.append(cell.value.strip() if isinstance(cell.value, str) else "")
            yield number, tuple(values + [""] * (len(headers) - len(values)))
    finally:
        rows.close()


def _identifier(value: str) -> bool:
    return re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,39}", value) is not None


def prepare_book_setup(
    book_id: str, label: str, word: str, currency_code: str, text_number: bool,
    master_filename: str, master: bytes, template_filename: str, template: bytes, *, version: int = 1,
    existing_profiles: Mapping[str, dict] | None = None,
) -> dict:
    if not _identifier(book_id) or not label or len(label) > 100:
        raise StatementError("请填写账套名称，账套标识仅用 1–40 位字母、数字、下划线或连字符，须以字母或数字开始。")
    if not word or len(word) > 20 or not currency_code or len(currency_code) > 20:
        raise StatementError("请填写凭证字和人民币币别代码，各不超过 20 个字符。")
    if any(Path(name).suffix.lower() != ".xlsx" for name in (master_filename, template_filename)):
        raise StatementError("主数据工作簿和精斗云空白模板都须为 .xlsx 文件。")

    subjects, auxiliary, accounts = [], [], []
    subject_requirements: dict[str, list[str]] = {}
    bank_eligible: set[str] = set()
    auxiliary_keys: set[tuple[str, str]] = set()
    account_ids: set[str] = set()
    numbers: set[str] = set()
    bank_subjects: dict[str, str] = {}
    with open_statement(master_filename, master, SETUP_LIMITS) as reader:
        for row, (code, name, categories, quantity) in _rows(reader, SUBJECTS):
            required = [category.strip() for category in re.split(r"[,，、]", categories)] if categories else []
            if (not code or len(code) > 80 or not name or code in subject_requirements or quantity not in ("", "是", "否")
                    or "" in required or len(set(required)) != len(required) or len(set(required) - set(STANDARD_AUX)) > 2):
                raise StatementError(f"「科目」第 {row} 行有空值、重复编码或无效的辅助类别、数量核算选项。")
            subjects.append({"code": code, "name": name, "required_aux": required, "quantity_required": quantity == "是"})
            subject_requirements[code] = required
            if not required and quantity != "是":
                bank_eligible.add(code)
        for row, (category, code) in _rows(reader, AUXILIARY):
            if not category or not code or (category, code) in auxiliary_keys:
                raise StatementError(f"「辅助核算」第 {row} 行有空值，或辅助类别与编码组合重复。")
            auxiliary.append({"category": category, "code": code})
            auxiliary_keys.add((category, code))
        for row, (account_id, name, number, subject, bank_format, date_preference) in _rows(reader, ACCOUNTS):
            if (not _identifier(account_id) or account_id in account_ids or not name
                    or not re.fullmatch(r"[0-9]{1,40}", number) or number in numbers):
                raise StatementError(f"「银行账户」第 {row} 行的账户标识、名称或完整账号缺失、无效或重复。")
            if subject not in bank_eligible:
                raise StatementError(f"「银行账户」第 {row} 行须使用「科目」表中不要求辅助或数量核算的银行科目。")
            if bank_format not in ("", "自动", "自定义", *FORMATS) or date_preference not in ("", *DATE_PREFERENCES):
                raise StatementError(f"「银行账户」第 {row} 行的流水格式或日期口径无效，请按页面说明填写。")
            item: dict = {
                "id": account_id, "label": name, "number": number, "profile_id": FORMATS.get(bank_format),
                "posting_preference": DATE_PREFERENCES.get(date_preference),
            }
            if bank_format == "自定义":
                profile = existing_profiles.get(account_id) if existing_profiles is not None else None
                if profile is None:
                    raise StatementError(f"「银行账户」第 {row} 行没有已发布的自定义格式，请先使用自动或内置格式。")
                item["bank_profile"] = profile
            accounts.append(item)
            account_ids.add(account_id)
            numbers.add(number)
            bank_subjects[account_id] = subject
    if not subjects or not accounts:
        raise StatementError("至少要填写一个科目和一个银行账户。")
    categories = {category for category, _ in auxiliary_keys}
    for code, required in subject_requirements.items():
        if set(required) - categories:
            raise StatementError(f"科目「{code}」要求的辅助类别在「辅助核算」表中没有编码。")
    validate_empty_template(template, SETUP_LIMITS)
    return {
        "ledger": {
            "book_id": book_id, "version": version, "bank_subjects": bank_subjects,
            "subjects": subjects, "auxiliary": auxiliary, "finance_verified": True,
        },
        "rules": [],
        "web": {
            "label": label, "accounts": accounts, "verified_template_hash": sha256(template).hexdigest(),
            "voucher_word": word, "currency_code": currency_code, "text_number": text_number,
        },
    }
