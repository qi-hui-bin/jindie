from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from datetime import date, time
from typing import Literal, Mapping

Direction = Literal["income", "expense"]
RawValue = str | int | float | bool | None


def canonical_json(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def fingerprint(value: object) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class BookAccount:
    book_id: str
    account_id: str
    number: str

    def __post_init__(self) -> None:
        if not self.book_id.strip() or not self.account_id.strip() or not re.fullmatch(r"[0-9]+", self.number):
            raise ValueError("请提供账套、银行账户标识及可核对的完整数字账号。")


@dataclass(frozen=True)
class Period:
    year: int
    month: int

    def __post_init__(self) -> None:
        date(self.year, self.month, 1)

    @classmethod
    def from_text(cls, value: str) -> Period:
        if not re.fullmatch(r"[0-9]{4}-[0-9]{2}", value):
            raise ValueError("会计期间格式应为 YYYY-MM。")
        return cls(int(value[:4]), int(value[5:]))

    def __str__(self) -> str:
        return f"{self.year:04d}-{self.month:02d}"

    def contains(self, value: date) -> bool:
        return (self.year, self.month) == (value.year, value.month)


@dataclass(frozen=True)
class StatementTransaction:
    bank_reference: str | None
    own_account: str
    counterparty_account: str | None
    counterparty_name: str | None
    transaction_date: date
    transaction_time: time | None
    posting_date: date | None
    voucher_date: date
    direction: Direction
    amount_minor: int
    currency: str
    balance_minor: int | None
    summary: str | None
    purpose: str | None
    message: str | None
    remark: str | None

    def key_fields(self) -> dict[str, RawValue]:
        return {
            "transaction_date": self.transaction_date.isoformat(),
            "transaction_time": self.transaction_time.isoformat() if self.transaction_time else None,
            "direction": self.direction,
            "amount_minor": self.amount_minor,
            "currency": self.currency,
            "counterparty_account": self.counterparty_account,
            "counterparty_name": self.counterparty_name,
        }

    def payload(self) -> dict[str, RawValue]:
        return {
            **self.key_fields(),
            "bank_reference": self.bank_reference,
            "own_account": self.own_account,
            "posting_date": self.posting_date.isoformat() if self.posting_date else None,
            "voucher_date": self.voucher_date.isoformat(),
            "balance_minor": self.balance_minor,
            "summary": self.summary,
            "purpose": self.purpose,
            "message": self.message,
            "remark": self.remark,
        }


@dataclass(frozen=True)
class Issue:
    code: str
    message: str


@dataclass(frozen=True)
class StatementRow:
    row_number: int
    raw: Mapping[str, RawValue]
    issues: tuple[Issue, ...]
    transaction: StatementTransaction | None = None


@dataclass(frozen=True)
class ParsedStatement:
    filename: str
    file_hash: str
    sheet: str
    header_row: int
    profile_id: str
    profile_version: int
    profile_json: str
    account: BookAccount
    period: Period
    posting_preference: Literal["transaction", "posting"] | None
    confirmed_binding: bool
    warnings: tuple[Issue, ...]
    rows: tuple[StatementRow, ...]
