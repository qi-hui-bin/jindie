from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass

from .accounting import AccountingError, Assignment, Condition, Exclusion, Ledger, Rule, Voucher, propose
from .storage import StatementStore
from .voucher_preview import batch_sources

MATCH_FIELDS = {"summary": "银行摘要", "purpose": "用途", "message": "交易附言", "remark": "备注"}


def rule_config(
    book_id: str, values: Mapping[str, str], assignment: Assignment, version: int,
) -> dict:
    rule_id, party, field, text = (values[key] for key in ("rule_id", "party", "field", "text"))
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,39}", rule_id):
        raise AccountingError("规则标识须为 1–40 位字母、数字、下划线或连字符，以字母或数字开头。")
    if len(party) > 200 or len(text) > 200:
        raise AccountingError("对方户名和匹配文本各不得超过 200 个字符。")
    if (not party and not text) or bool(field) != bool(text):
        raise AccountingError("请填写精确对方户名或完整的字段匹配条件，不能仅按收支方向匹配。")
    conditions = []
    if field:
        if field not in MATCH_FIELDS:
            raise AccountingError("请选择可匹配的银行字段。")
        operation = values["operation_type"]
        if operation not in ("equals", "contains"):
            raise AccountingError("请选择精确等于或包含。")
        condition = Condition(field, operation, (text,))
        conditions.append({"field": condition.field, "operation": condition.operation, "values": list(condition.values)})
    return {
        "id": rule_id, "version": version, "book_id": book_id, "account_id": values["account_id"],
        "direction": values["direction"], "counterpart_names": [party] if party else [], "conditions": conditions,
        "counterpart": {"subject": assignment.subject, "auxiliary": [
            {"category": aux.category, "code": aux.code} for aux in assignment.auxiliary
        ]},
        "summary": values["summary"], "enabled": True,
    }


def validate_active_rules(ledger: Ledger, rules: tuple[Rule, ...]) -> None:
    enabled: set[str] = set()
    for rule in rules:
        if not rule.enabled:
            continue
        if rule.book_id != ledger.book_id or rule.account_id not in ledger.bank_subjects or rule.id in enabled:
            raise AccountingError(f"启用规则「{rule.id}」的账套、账户或标识无效；请先停用规则再调整账户。")
        enabled.add(rule.id)
        targets = (rule.counterpart,) if rule.counterpart is not None else tuple(line.assignment for line in rule.fixed_entries)
        for target in targets:
            if issues := ledger.validate(ledger.resolve(target)):
                raise AccountingError(f"启用规则「{rule.id}」：" + "；".join(issue.message for issue in issues))


@dataclass(frozen=True)
class RuleTrial:
    matched: int
    won: int
    sample: tuple[Voucher, ...]
    replaced: tuple[str, ...]
    issues: tuple[Exclusion, ...]
    shadowed: tuple[int, ...]

    @property
    def can_publish(self) -> bool:
        return self.matched > 0 and self.matched == self.won


def trial_rule(
    store: StatementStore, batch_id: int, ledger: Ledger, current: tuple[Rule, ...], proposed: tuple[Rule, ...],
    candidate: Rule,
) -> RuleTrial:
    _, sources, _ = batch_sources(store, batch_id, ledger, account_id=candidate.account_id)
    matched = {source.row_number for source in sources if candidate.matches(source)}
    proposal = propose(sources, ledger, proposed)
    won = tuple(voucher for voucher in proposal.vouchers if (candidate.id, candidate.version) in voucher.rule_versions)
    won_rows = {source.row_number for voucher in won for source in voucher.sources}
    issues = tuple(item for item in proposal.exclusions if item.row_number in matched)
    shadowed = tuple(sorted(matched - won_rows - {item.row_number for item in issues}))
    baseline = propose(sources, ledger, current)
    replaced = {rule_id for voucher in baseline.vouchers for rule_id, _ in voucher.rule_versions
                if rule_id != candidate.id and any(source.row_number in won_rows for source in voucher.sources)}
    return RuleTrial(len(matched), len(won_rows), won[:12], tuple(sorted(replaced)), issues[:12], shadowed[:12])
