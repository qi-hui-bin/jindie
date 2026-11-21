from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass, replace
from datetime import date, time
from string import Formatter
from typing import Literal, Mapping

from .models import Direction, Issue, Period, StatementTransaction, fingerprint

MergeMode = Literal["none", "bank_detail", "summarize"]
Side = Literal["debit", "credit"]
STANDARD_AUX = ("客户", "供应商", "职员", "项目", "部门", "存货")
TEXT_FIELDS = ("summary", "purpose", "message", "remark", "counterparty_account")
MAX_MINOR = 999_999_999_999_999


class AccountingError(ValueError):
    """A configuration or review action violates a voucher invariant."""


@dataclass(frozen=True, order=True)
class Auxiliary:
    category: str
    code: str

    def __post_init__(self) -> None:
        if not isinstance(self.category, str) or not isinstance(self.code, str) or not self.category.strip() or not self.code.strip():
            raise AccountingError("辅助类别和编码必须是非空文本。")


@dataclass(frozen=True)
class Subject:
    code: str
    name: str
    required_aux: tuple[str, ...] = ()
    quantity_required: bool = False


@dataclass(frozen=True)
class Assignment:
    subject: str
    auxiliary: tuple[Auxiliary, ...] = ()


@dataclass(frozen=True)
class Ledger:
    book_id: str
    version: int
    bank_subjects: Mapping[str, str]
    subjects: Mapping[str, Subject]
    auxiliary: frozenset[Auxiliary] = frozenset()
    legacy_assignments: Mapping[str, Assignment] | None = None
    finance_verified: bool = False

    def __post_init__(self) -> None:
        if not self.book_id.strip() or self.version < 1 or not self.bank_subjects:
            raise AccountingError("请设置账套、主数据版本和银行科目。")
        if type(self.finance_verified) is not bool:
            raise AccountingError("主数据核验状态必须是布尔值。")
        if any(code != item.code or not code.strip() or not item.name.strip() for code, item in self.subjects.items()):
            raise AccountingError("科目编码必须与索引一致，且编码和名称均不能为空。")
        if any(code not in self.subjects for code in self.bank_subjects.values()):
            raise AccountingError("银行账户配置引用了不存在的科目。")

    def resolve(self, assignment: Assignment) -> Assignment:
        assignment = ((self.legacy_assignments or {}).get(assignment.subject, assignment)
                      if not assignment.auxiliary else assignment)
        return Assignment(assignment.subject, tuple(sorted(assignment.auxiliary)))

    def validate(self, assignment: Assignment) -> tuple[Issue, ...]:
        subject = self.subjects.get(assignment.subject)
        if subject is None:
            return (Issue("unknown_subject", f"科目「{assignment.subject}」不在本账套主数据中。"),)
        if subject.quantity_required:
            return (Issue("quantity_required", f"科目「{subject.code}」需要数量核算，请财务复核。"),)
        supplied = {aux.category for aux in assignment.auxiliary}
        required = set(subject.required_aux)
        problems = [
            Issue("missing_auxiliary", f"科目「{subject.code}」缺少「{category}」编码。")
            for category in sorted(required - supplied)
        ]
        if supplied - required or len(supplied) != len(assignment.auxiliary) or len(supplied - set(STANDARD_AUX)) > 2:
            problems.append(Issue("invalid_auxiliary", "辅助类别不适用、重复或超出模板的自定义列数。"))
        if any(aux not in self.auxiliary for aux in assignment.auxiliary):
            problems.append(Issue("unknown_auxiliary", "辅助类别和编码的组合不在本账套主数据中。"))
        return tuple(problems)


@dataclass(frozen=True)
class Condition:
    field: str
    operation: Literal["equals", "contains"]
    values: tuple[str, ...]

    def __post_init__(self) -> None:
        if self.field not in TEXT_FIELDS or self.operation not in ("equals", "contains") or (
            not self.values or any(not value.strip() for value in self.values)
        ):
            raise AccountingError("匹配条件需要有效字段、方式和非空关键词。")

    def matches(self, transaction: StatementTransaction) -> bool:
        text = (getattr(transaction, self.field) or "").strip().casefold()
        values = (value.strip().casefold() for value in self.values)
        return any((text == value if self.operation == "equals" else value in text) for value in values)


@dataclass(frozen=True)
class FixedEntry:
    assignment: Assignment
    side: Side
    amount_minor: int

    def __post_init__(self) -> None:
        if self.side not in ("debit", "credit") or type(self.amount_minor) is not int or not 0 < self.amount_minor <= MAX_MINOR:
            raise AccountingError("固定分录须提供借贷方向及正整数分。")


@dataclass(frozen=True)
class Rule:
    id: str
    version: int
    book_id: str
    direction: Direction
    counterpart: Assignment | None
    summary: str
    account_id: str | None = None
    counterpart_names: tuple[str, ...] = ()
    conditions: tuple[Condition, ...] = ()
    priority: int | None = None
    merge: MergeMode = "none"
    enabled: bool = False
    fixed_entries: tuple[FixedEntry, ...] = ()
    bank_references: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not self.id.strip() or type(self.version) is not int or self.version < 1 or not self.book_id.strip() or (
            self.direction not in ("income", "expense") or self.merge not in ("none", "bank_detail", "summarize")
            or not self.summary.strip() or any(not name.strip() for name in self.counterpart_names)
            or type(self.enabled) is not bool or (self.priority is not None and type(self.priority) is not int)
        ):
            raise AccountingError("凭证规则的标识、适用范围、摘要或合并方式无效。")
        if (self.counterpart is None) != bool(self.fixed_entries):
            raise AccountingError("规则须选择普通对方分录或固定多分录。")
        if self.fixed_entries or self.bank_references:
            if not (self.fixed_entries and self.account_id and self.bank_references and self.merge == "bank_detail"
                    and len(set(self.bank_references)) == len(self.bank_references)
                    and all(isinstance(ref, str) and ref.strip() for ref in self.bank_references)):
                raise AccountingError("固定多分录须限定账户和唯一流水号，且使用 bank_detail 模式。")
        try:
            if any(
                field is not None and (field not in (*TEXT_FIELDS, "counterparty", "date") or spec or conversion)
                for _, field, spec, conversion in Formatter().parse(self.summary)
            ):
                raise ValueError
        except ValueError as exc:
            raise AccountingError(f"规则「{self.id}」的摘要格式无效。") from exc

    @property
    def rank(self) -> int:
        return self.priority if self.priority is not None else (
            100 * bool(self.fixed_entries) + 20 * bool(self.conditions) + 10 * bool(self.counterpart_names)
        )

    def matches(self, source: Source) -> bool:
        tx = source.transaction
        other = (tx.counterparty_name or "").strip().casefold()
        return (
            self.enabled and self.book_id == source.book_id and self.account_id in (None, source.account_id)
            and tx.direction == self.direction
            and (not self.bank_references or tx.bank_reference in self.bank_references)
            and (not self.counterpart_names or other in {name.strip().casefold() for name in self.counterpart_names})
            and all(condition.matches(tx) for condition in self.conditions)
        )

    def render_summary(self, tx: StatementTransaction) -> str:
        substitutions = {field: getattr(tx, field) or "" for field in TEXT_FIELDS}
        substitutions["counterparty"] = tx.counterparty_name or ""
        substitutions["date"] = tx.voucher_date.isoformat()
        return self.summary.format_map(substitutions).strip()


@dataclass(frozen=True)
class Source:
    transaction_id: int
    book_id: str
    account_id: str
    period: Period
    batch_id: int
    sheet: str
    row_number: int
    transaction: StatementTransaction

    @property
    def order(self) -> tuple[date, time, int, int]:
        return self.transaction.voucher_date, self.transaction.transaction_time or time.min, self.row_number, self.transaction_id


@dataclass(frozen=True)
class Entry:
    assignment: Assignment
    summary: str
    side: Side
    amount_minor: int
    transaction_ids: tuple[int, ...]


@dataclass(frozen=True)
class Voucher:
    sources: tuple[Source, ...]
    entries: tuple[Entry, ...]
    rule_versions: tuple[tuple[str, int], ...] = ()
    review_digest: str | None = None

    @property
    def order(self) -> tuple[date, time, int, int]:
        return min(source.order for source in self.sources)

    def digest(self, ledger: Ledger) -> str:
        return fingerprint({
            "ledger": [
                ledger.book_id, ledger.version,
                sorted(ledger.bank_subjects.items()),
                [asdict(ledger.subjects[code]) for code in sorted({entry.assignment.subject for entry in self.entries})
                 if code in ledger.subjects],
            ],
            "sources": [
                [s.transaction_id, s.book_id, s.account_id, str(s.period), s.batch_id, s.sheet, s.row_number, s.transaction.payload()]
                for s in self.sources
            ],
            "entries": [asdict(entry) for entry in self.entries],
            "rules": self.rule_versions,
        })


@dataclass(frozen=True)
class Exclusion:
    row_number: int
    issues: tuple[Issue, ...]
    skipped: bool = False


@dataclass(frozen=True)
class Proposal:
    vouchers: tuple[Voucher, ...]
    exclusions: tuple[Exclusion, ...]


def validate_voucher(voucher: Voucher, ledger: Ledger) -> tuple[Issue, ...]:
    if not voucher.sources:
        return (Issue("missing_sources", "凭证没有来源流水。"),)
    first = voucher.sources[0]
    bank = ledger.bank_subjects.get(first.account_id)
    sources = {source.transaction_id: source for source in voucher.sources}
    scope = {(s.book_id, s.account_id, s.period, s.transaction.voucher_date, s.transaction.currency) for s in voucher.sources}
    if len(sources) != len(voucher.sources) or len(scope) != 1 or first.book_id != ledger.book_id or not bank:
        return (Issue("invalid_sources", "同张凭证的流水必须唯一，并属于同一账套、账户、期间、日期和币种。"),)
    issues: list[Issue] = []
    if first.transaction.currency != "CNY":
        issues.append(Issue("foreign_currency", "非人民币流水请交财务复核。"))
    if not first.period.contains(first.transaction.voucher_date):
        issues.append(Issue("out_of_period", "凭证日期不在所选期间。"))
    totals = {"debit": 0, "credit": 0}
    bank_refs = {"debit": set(), "credit": set()}
    for entry in voucher.entries:
        issues.extend(ledger.validate(entry.assignment))
        refs = set(entry.transaction_ids)
        if not entry.summary.strip():
            issues.append(Issue("empty_summary", "每条分录都必须填写摘要。"))
        if entry.side not in totals or type(entry.amount_minor) is not int or not 0 < entry.amount_minor <= MAX_MINOR:
            issues.append(Issue("invalid_entry", "分录必须只有一侧且金额为可导出的正整数分。"))
            continue
        totals[entry.side] += entry.amount_minor
        if not refs or len(refs) != len(entry.transaction_ids) or not refs <= sources.keys():
            issues.append(Issue("invalid_entry_sources", "分录缺少来源或关联了凭证之外的流水。"))
            continue
        if entry.assignment.subject == bank:
            direction = "income" if entry.side == "debit" else "expense"
            if refs & bank_refs[entry.side] or any(sources[ref].transaction.direction != direction for ref in refs) or (
                sum(sources[ref].transaction.amount_minor for ref in refs) != entry.amount_minor
            ):
                issues.append(Issue("bank_mismatch", "银行科目分录的方向、金额或来源与流水不符。"))
            bank_refs[entry.side].update(refs)
    if len(voucher.entries) < 2 or totals["debit"] != totals["credit"]:
        issues.append(Issue("unbalanced", f"凭证借贷相差 {totals['debit'] - totals['credit']} 分。"))
    if bank_refs["debit"] | bank_refs["credit"] != sources.keys():
        issues.append(Issue("bank_mismatch", "有来源流水未计入银行分录。"))
    return tuple(issues)


def review(voucher: Voucher, ledger: Ledger) -> Voucher:
    if issues := validate_voucher(voucher, ledger):
        raise AccountingError("；".join(issue.message for issue in issues))
    return replace(voucher, review_digest=voucher.digest(ledger))


def edit_entries(voucher: Voucher, entries: tuple[Entry, ...], ledger: Ledger) -> tuple[Voucher, tuple[Issue, ...]]:
    edited = replace(voucher, entries=entries, review_digest=None)
    return edited, validate_voucher(edited, ledger)


def _checked_draft(voucher: Voucher, ledger: Ledger) -> Voucher:
    if issues := validate_voucher(voucher, ledger):
        raise AccountingError("；".join(issue.message for issue in issues))
    return voucher


def manual_voucher(source: Source, ledger: Ledger, counterpart: Assignment, summary: str) -> Voucher:
    bank = ledger.bank_subjects.get(source.account_id)
    if not bank:
        raise AccountingError("本账套未配置该账户的银行科目。")
    counterpart = ledger.resolve(counterpart)
    if counterpart.subject == bank:
        raise AccountingError("对方科目不能与本方银行科目相同。")
    tx = source.transaction
    if tx.direction not in ("income", "expense"):
        raise AccountingError("来源流水的收支方向无效。")
    income = tx.direction == "income"
    refs = (source.transaction_id,)
    bank_entry = Entry(Assignment(bank), summary, "debit" if income else "credit", tx.amount_minor, refs)
    other_entry = Entry(counterpart, summary, "credit" if income else "debit", tx.amount_minor, refs)
    entries = (bank_entry, other_entry) if income else (other_entry, bank_entry)
    return _checked_draft(Voucher((source,), entries), ledger)


def merge_vouchers(vouchers: tuple[Voucher, ...], ledger: Ledger) -> Voucher:
    if len(vouchers) < 2:
        raise AccountingError("请选择至少两张凭证。")
    ordered = sorted(vouchers, key=lambda item: item.order)
    return _checked_draft(Voucher(
        tuple(sorted((s for item in ordered for s in item.sources), key=lambda s: s.order)),
        tuple(e for item in ordered for e in item.entries),
        tuple(sorted({rule for item in ordered for rule in item.rule_versions})),
    ), ledger)


def split_voucher(voucher: Voucher, ledger: Ledger) -> tuple[Voucher, ...]:
    if len(voucher.sources) < 2:
        raise AccountingError("至少两笔来源流水才能拆分。")
    _checked_draft(voucher, ledger)
    amounts = {source.transaction_id: source.transaction.amount_minor for source in voucher.sources}
    by_source: dict[int, list[Entry]] = defaultdict(list)
    for entry in voucher.entries:
        refs = entry.transaction_ids
        if len(refs) > 1 and entry.amount_minor != sum(amounts[ref] for ref in refs):
            raise AccountingError("多来源分录金额无法按流水拆分，请先核对分录。")
        for ref in refs:
            by_source[ref].append(replace(entry, amount_minor=amounts[ref] if len(refs) > 1 else entry.amount_minor,
                                          transaction_ids=(ref,)))
    return tuple(_checked_draft(Voucher((source,), tuple(by_source[source.transaction_id]), voucher.rule_versions), ledger)
                 for source in sorted(voucher.sources, key=lambda item: item.order))


def propose(
    sources: tuple[Source, ...],
    ledger: Ledger,
    rules: tuple[Rule, ...],
    *,
    confirmed_skips: Mapping[str, str] | None = None,
) -> Proposal:
    if len({s.transaction_id for s in sources}) != len(sources):
        raise AccountingError("发现重复来源，请先处理查重结果。")
    if len({rule.id for rule in rules if rule.enabled and rule.book_id == ledger.book_id}) != sum(
        rule.enabled and rule.book_id == ledger.book_id for rule in rules
    ):
        raise AccountingError("同一账套的每个规则只能启用一个版本。")
    skips = confirmed_skips or {}
    available = {s.transaction.bank_reference for s in sources}
    if any(not key.strip() or key not in available or not reason.strip() for key, reason in skips.items()):
        raise AccountingError("跳过项目须对应当批流水号并填写已核实的原因。")
    if skips and len({(s.book_id, s.account_id, s.period, s.batch_id) for s in sources}) != 1:
        raise AccountingError("已核实的跳过项目只能应用于同一批次。")
    excluded: list[Exclusion] = []
    grouped: dict[tuple[object, ...], list[tuple[Source, Assignment | None, str, Rule]]] = defaultdict(list)
    for source in sorted(sources, key=lambda item: item.order):
        tx = source.transaction
        if source.book_id != ledger.book_id or source.account_id not in ledger.bank_subjects:
            raise AccountingError("来源流水和所选账套主数据或账户不一致。")
        if tx.bank_reference in skips:
            excluded.append(Exclusion(source.row_number, (Issue("confirmed_skip", skips[tx.bank_reference]),), True))
            continue
        if tx.currency != "CNY":
            excluded.append(Exclusion(source.row_number, (Issue("foreign_currency", "非人民币流水请交财务复核。"),)))
            continue
        candidates = [rule for rule in rules if rule.matches(source)]
        if not candidates:
            excluded.append(Exclusion(source.row_number, (Issue("unmatched", "无已启用的匹配规则，请人工复核。"),)))
            continue
        highest = max(rule.rank for rule in candidates)
        winners = sorted((r for r in candidates if r.rank == highest), key=lambda r: (r.id, r.version))
        actions = {
            (ledger.resolve(rule.counterpart) if rule.counterpart else None, rule.merge, rule.render_summary(tx),
             rule.fixed_entries, rule.bank_references)
            for rule in winners
        }
        if len(actions) != 1:
            names = "、".join(f"{rule.id} v{rule.version}" for rule in winners)
            excluded.append(Exclusion(source.row_number, (Issue("rule_conflict", f"以下同优先级规则结果不同：{names}。"),)))
            continue
        assignment, mode, summary, _, _ = next(iter(actions))
        rule = winners[0]
        assignments = (assignment,) if assignment is not None else tuple(
            ledger.resolve(entry.assignment) for entry in rule.fixed_entries
        )
        problems = ledger.validate(Assignment(ledger.bank_subjects[source.account_id]))
        problems += tuple(issue for target in assignments for issue in ledger.validate(target))
        if not summary:
            problems += (Issue("empty_summary", "规则生成的摘要为空，请人工复核。"),)
        if problems:
            excluded.append(Exclusion(source.row_number, problems))
            continue
        key = (source.book_id, source.account_id, source.period, tx.voucher_date, tx.currency, tx.direction,
               rule.id, rule.version, assignment, source.transaction_id if mode == "none" else None)
        grouped[key].append((source, assignment, summary, rule))
    vouchers: list[Voucher] = []
    for group in grouped.values():
        first, assignment, summary, rule = group[0]
        bank = Assignment(ledger.bank_subjects[first.account_id])
        refs = tuple(item[0].transaction_id for item in group)
        total = sum(item[0].transaction.amount_minor for item in group)
        income = first.transaction.direction == "income"
        bank_side: Side = "debit" if income else "credit"
        counterpart_side: Side = "credit" if income else "debit"
        if assignment is None and {item[0].transaction.bank_reference for item in group} != set(rule.bank_references):
            issue = Issue("incomplete_fixed_group", "固定多分录规则的流水号未在同批同日全部匹配，请核对原文件及跳过记录。")
            excluded.extend(Exclusion(item[0].row_number, (issue,)) for item in group)
            continue
        others = ((Entry(assignment, summary, counterpart_side, total, refs),) if assignment is not None else tuple(
            Entry(ledger.resolve(entry.assignment), summary, entry.side, entry.amount_minor, refs)
            for entry in rule.fixed_entries
        ))
        banks = ((Entry(bank, summary, bank_side, total, refs),) if rule.merge == "summarize" else tuple(
            Entry(bank, text, bank_side, source.transaction.amount_minor, (source.transaction_id,))
            for source, _, text, _ in group
        ))
        voucher = Voucher(tuple(item[0] for item in group), banks + others if income else others + banks,
                          ((rule.id, rule.version),))
        if issues := validate_voucher(voucher, ledger):
            excluded.extend(Exclusion(s.row_number, issues) for s in voucher.sources)
        else:
            vouchers.append(voucher)
    return Proposal(tuple(sorted(vouchers, key=lambda v: v.order)), tuple(sorted(excluded, key=lambda e: e.row_number)))
