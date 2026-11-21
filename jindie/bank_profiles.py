from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal, Mapping

from .models import Direction

AMOUNT_FIELDS = {
    "signed": frozenset({"amount"}),
    "coded_signed": frozenset({"amount", "direction_code"}),
    "split": frozenset({"income", "expense"}),
}
COLUMN_FIELDS = {
    "transaction_date", "transaction_time", "posting_date", "balance", "currency", "reference",
    "own_account", "counterparty_account", "counterparty_name", "payer_account", "payer_name",
    "payee_account", "payee_name", "summary", "purpose", "message", "remark", "value_date",
    "amount", "direction_code", "income", "expense",
}


def normalize_header(value: str, ignore_english_suffix: bool = False) -> str:
    value = value.replace("（", "(").replace("）", ")")
    if ignore_english_suffix and re.search(r"[\u4e00-\u9fff]", value):
        value = re.sub(r"\s*\[\s*[A-Za-z][^\]]*\]\s*$", "", value)
        value = re.sub(r"(?:[\s/:：-]+[A-Za-z][A-Za-z0-9 /.,_()-]*|\([A-Za-z][A-Za-z0-9 /.,_()-]*\))\s*$", "", value)
    return re.sub(r"\s+", "", value).casefold()


@dataclass(frozen=True)
class BankProfile:
    id: str
    version: int
    name: str
    columns: Mapping[str, tuple[str, ...]]
    amount_mode: Literal["signed", "coded_signed", "split"]
    currency_codes: Mapping[str, str]
    direction_codes: Mapping[str, Direction] = field(default_factory=dict)
    identify_by: tuple[str, ...] = ()
    sheet_name: str | None = None
    header_row: int | None = None
    header_scan_rows: int = 40
    account_label: str | None = None
    account_label_row: int | None = None
    default_currency: str | None = None
    allow_serial_dates: bool = False
    ignore_english_suffix: bool = False
    ignored_row_labels: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not self.id.strip() or not self.name.strip() or self.version < 1:
            raise ValueError("银行配置必须有标识、名称及正整数版本。")
        if not 1 <= self.header_scan_rows <= 1000 or self.header_row is not None and not 1 <= self.header_row <= 1000:
            raise ValueError("银行配置的表头位置或扫描范围无效。")
        if self.account_label_row is not None and not 1 <= self.account_label_row <= 1000:
            raise ValueError("银行配置的账号标签行无效。")
        if self.account_label_row is not None and not self.account_label:
            raise ValueError("指定账号标签行时还须提供账号标签。")
        if self.amount_mode not in AMOUNT_FIELDS or not self.required_fields <= self.columns.keys():
            raise ValueError("银行配置缺少必填字段，或收支模式无效。")
        if self.columns.keys() - COLUMN_FIELDS:
            raise ValueError("银行配置包含未知的标准字段。")
        if self.amount_mode == "coded_signed" and (
            not self.direction_codes or set(self.direction_codes.values()) - {"income", "expense"}
        ):
            raise ValueError("带交易类型的有符号金额模式须配置有效的收支码。")
        aliases: dict[str, str] = {}
        for column, values in self.columns.items():
            if not values or isinstance(values, str):
                raise ValueError("字段别名应是非空列表。")
            for value in values:
                alias = normalize_header(value, self.ignore_english_suffix)
                if not alias or (alias in aliases and aliases[alias] != column):
                    raise ValueError("表头别名为空或同时映射到多个字段。")
                aliases[alias] = column
        if any(not normalize_header(label, self.ignore_english_suffix) for label in self.ignored_row_labels):
            raise ValueError("要忽略的标题或页脚不得为空。")

    @property
    def required_fields(self) -> frozenset[str]:
        return AMOUNT_FIELDS[self.amount_mode] | {"transaction_date", *self.identify_by}


BOC_9130 = BankProfile(
    id="boc-9130",
    version=2,
    name="坤昇中行 9130",
    amount_mode="coded_signed",
    direction_codes={"来账": "income", "往账": "expense"},
    currency_codes={"CNY": "CNY"},
    identify_by=("currency",),
    account_label="查询账号",
    account_label_row=2,
    ignore_english_suffix=True,
    ignored_row_labels=("合计", "总计"),
    columns={
        "direction_code": ("交易类型",),
        "amount": ("交易金额",),
        "transaction_date": ("交易日期",),
        "transaction_time": ("交易时间",),
        "value_date": ("起息日期",),
        "payer_name": ("付款人名称", "付款人户名"),
        "payer_account": ("付款人账号",),
        "payee_name": ("收款人名称", "收款人户名"),
        "payee_account": ("收款人账号",),
        "summary": ("摘要",),
        "purpose": ("用途",),
        "message": ("交易附言",),
        "remark": ("备注",),
        "reference": ("交易流水号",),
        "balance": ("交易后余额",),
        "currency": ("交易货币",),
    },
)

CCB_0306 = BankProfile(
    id="ccb-0306",
    version=2,
    name="文烁建行 0306",
    amount_mode="split",
    currency_codes={"人民币元": "CNY"},
    identify_by=("posting_date", "own_account", "currency"),
    ignored_row_labels=("合计", "总计"),
    columns={
        "own_account": ("账号",),
        "expense": ("借方发生额（支取）",),
        "income": ("贷方发生额（收入）",),
        "transaction_date": ("交易时间",),
        "posting_date": ("记账日期",),
        "counterparty_name": ("对方户名",),
        "counterparty_account": ("对方账号",),
        "summary": ("摘要",),
        "remark": ("备注",),
        "reference": ("账户明细编号-交易流水号",),
        "balance": ("余额",),
        "currency": ("币种",),
    },
)

CIB = BankProfile(
    id="cib",
    version=1,
    name="兴业银行",
    amount_mode="split",
    currency_codes={"人民币": "CNY", "CNY": "CNY"},
    identify_by=("own_account", "currency", "reference"),
    ignored_row_labels=("合计", "总计"),
    columns={
        "own_account": ("账号",),
        "expense": ("借方金额(支出)",),
        "income": ("贷方金额(收入)",),
        "transaction_date": ("记账日期",),
        "transaction_time": ("交易时间",),
        "counterparty_name": ("对方户名",),
        "counterparty_account": ("对方账号",),
        "summary": ("摘要",),
        "purpose": ("用途",),
        "remark": ("备注",),
        "reference": ("唯一流水编号",),
        "balance": ("账户余额",),
        "currency": ("币种",),
    },
)

BUILTIN_PROFILES = (BOC_9130, CCB_0306, CIB)
