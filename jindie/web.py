from __future__ import annotations

import json
import re
import secrets
import sqlite3
from dataclasses import asdict, dataclass
from datetime import date
from io import BytesIO
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, session, url_for
from itsdangerous import BadData, URLSafeTimedSerializer
from werkzeug.exceptions import HTTPException

from .accounting import AccountingError, Assignment, Auxiliary, validate_voucher
from .bank_import import PostingPreference, parse_statement
from .bank_profiles import BUILTIN_PROFILES, BankProfile
from .bank_setup import COLUMNS, custom_profile, mapping_preview
from .book_setup import SETUP_LIMITS, prepare_book_setup, setup_workbook
from .excel_reader import ImportLimits
from .kingdee_export import ExportOptions
from .models import BookAccount, Period
from .rule_setup import MATCH_FIELDS, rule_config, trial_rule, validate_active_rules
from .storage import StaleBookSetup
from .voucher_preview import parse_config
from .voucher_workflow import (
    PENDING_GROUPS, SKIP_REASONS, StaleReview, VoucherWorkflow, can_resolve_manually, review_exclusions,
)

LIMITS = ImportLimits()
UPLOAD_LIMIT = LIMITS.file_bytes + 1024 * 1024
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
IMPORT_LABELS = {"generated": "已生成 Excel", "imported": "已确认导入", "failed": "导入失败", "not_imported": "已核实未导入"}


@dataclass(frozen=True)
class ConfiguredAccount:
    account: BookAccount
    label: str
    posting_preference: PostingPreference | None
    profiles: tuple[BankProfile, ...] = BUILTIN_PROFILES


@dataclass(frozen=True)
class BookSettings:
    id: str
    label: str
    config: dict
    accounts: dict[str, ConfiguredAccount]
    template: Path | None
    template_hash: str | None
    word: str
    currency_code: str
    text_number: bool
    verified: bool
    managed: bool = False

    @property
    def version(self) -> int:
        return self.config["ledger"]["version"]

    @property
    def can_export(self) -> bool:
        return self.verified and (self.managed or self.template is not None) and self.template_hash is not None


def load_book_settings(path: Path) -> BookSettings:
    return _book_settings(json.loads(path.read_text(encoding="utf-8-sig")), config_dir=path.parent)


def _book_settings(data: dict, *, config_dir: Path | None = None) -> BookSettings:
    try:
        ledger, rules = parse_config(data)
        if config_dir is None:
            validate_active_rules(ledger, rules)
        settings = data["web"]
        accounts: dict[str, ConfiguredAccount] = {}
        for item in settings["accounts"]:
            account = BookAccount(ledger.book_id, item["id"], item["number"])
            preference = item.get("posting_preference")
            if (account.account_id in accounts or account.account_id not in ledger.bank_subjects
                    or preference not in (None, "transaction", "posting")):
                raise AccountingError("本地账户配置无效、重复或缺少银行科目。")
            profile_id = item.get("profile_id")
            if "bank_profile" in item:
                profiles = (BankProfile(**item["bank_profile"]),)
            elif profile_id is not None:
                profiles = tuple(profile for profile in BUILTIN_PROFILES if profile.id == profile_id)
                if not profiles:
                    raise AccountingError("银行账户引用了未知的内置流水格式。")
            else:
                profiles = BUILTIN_PROFILES
            accounts[account.account_id] = ConfiguredAccount(
                account, item.get("label", account.account_id), preference, profiles,
            )
        if not accounts or len({item.account.number for item in accounts.values()}) != len(accounts):
            raise AccountingError("请配置至少一个银行账户，且完整账号不得重复。")
        template = config_dir / settings["template"] if config_dir is not None and settings.get("template") else None
        template_hash = settings.get("verified_template_hash")
        if template_hash is not None and not re.fullmatch(r"[a-f0-9]{64}", template_hash):
            raise AccountingError("模板核验值须为 64 位小写 SHA-256。")
        word, currency = settings.get("voucher_word", "记"), settings.get("currency_code", "RMB")
        text_number = settings.get("text_number", False)
        if not word.strip() or not currency.strip() or type(text_number) is not bool:
            raise AccountingError("凭证字、币别或编号类型配置无效。")
        # Freeze accounting inputs without storing deployment paths or account settings in review snapshots.
        return BookSettings(ledger.book_id, settings.get("label", ledger.book_id),
                            {"ledger": data["ledger"], "rules": data["rules"]}, accounts, template,
                            template_hash, word, currency, text_number, ledger.finance_verified, config_dir is None)
    except (TypeError, KeyError, AttributeError) as exc:
        raise AccountingError("本地账套配置缺少有效的 web 设置，请核对操作说明。") from exc


def _csrf() -> str:
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


def _positive(value: str, label: str) -> int:
    if not re.fullmatch(r"[0-9]{1,15}", value) or int(value) < 1:
        raise ValueError(f"「{label}」须为正整数（最多 15 位）。")
    return int(value)


def _revision() -> int:
    text = request.form.get("revision", "")
    if not re.fullmatch(r"[0-9]{1,12}", text):
        abort(400)
    return int(text)


def _entry_fields() -> tuple[str, Assignment]:
    summary, subject, text = (request.form.get(name, "").strip() for name in ("summary", "subject", "aux"))
    if not summary or not subject or len(summary) > 500 or len(subject) > 80 or len(text) > 500:
        raise AccountingError("请填写有效的摘要和科目编码。")
    pairs = tuple(value.strip().partition("=") for value in text.splitlines())
    if any(not category.strip() or not equals or not code.strip() for category, equals, code in pairs):
        raise AccountingError("辅助核算请每行填写「类别=编码」，也可以留空。")
    return summary, Assignment(subject, tuple(Auxiliary(category.strip(), code.strip()) for category, _, code in pairs))


def _money(value: int) -> str:
    sign = "-" if value < 0 else ""
    major, minor = divmod(abs(value), 100)
    return f"{sign}{major:,}.{minor:02d}"


def create_app(database: Path, configs: list[Path], *, allow_setup: bool = True) -> Flask:
    store = VoucherWorkflow(database)
    store.initialize()
    file_books: dict[str, BookSettings] = {}
    for config in configs:
        book = load_book_settings(config)
        if book.id in file_books:
            raise AccountingError("启动时不得提供同一账套的多个配置。")
        file_books[book.id] = book
    if not file_books and not allow_setup:
        raise AccountingError("启动时至少需要提供一份账套配置。")

    def configured_books() -> dict[str, BookSettings]:
        result = dict(file_books)
        for config in store.current_book_setups() if allow_setup else ():
            book = _book_settings(config)
            if book.id in result:
                raise AccountingError("数据库和本地 JSON 包含相同的账套标识，请移走重复的 JSON 配置。")
            result[book.id] = book
        return result

    configured_books()
    app = Flask(__name__)
    secret = secrets.token_hex(32)
    preview_signer = URLSafeTimedSerializer(secret, salt="bank-mapping-preview")
    rule_signer = URLSafeTimedSerializer(secret, salt="voucher-rule-preview")
    recalculation_signer = URLSafeTimedSerializer(secret, salt="pending-row-recalculation")
    app.config.update(SECRET_KEY=secret, TRUSTED_HOSTS=["127.0.0.1", "localhost"],
                      MAX_CONTENT_LENGTH=UPLOAD_LIMIT, MAX_FORM_PARTS=500,
                      SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Strict")
    app.jinja_env.globals.update(csrf_token=_csrf, import_labels=IMPORT_LABELS)
    app.jinja_env.filters["money"] = _money

    def get_book(book_id: str) -> BookSettings:
        book = configured_books().get(book_id)
        if book is None:
            abort(404)
        return book

    @app.before_request
    def check_request() -> None:
        if request.method == "POST":
            origin = request.headers.get("Origin")
            if origin and origin != request.host_url.rstrip("/"):
                abort(403)
            actual, expected = request.form.get("csrf", ""), session.get("csrf")
            if not expected or not secrets.compare_digest(actual.encode("utf-8"), expected.encode("utf-8")):
                abort(403)

    @app.after_request
    def protect(response):
        response.headers["Cache-Control"] = "no-store"
        response.headers["Content-Security-Policy"] = "default-src 'self'; frame-ancestors 'none'; form-action 'self'; base-uri 'none'"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        return response

    @app.errorhandler(ValueError)
    def validation_error(error: ValueError):
        return render_template("error.html", message=str(error)), 409 if isinstance(error, (StaleReview, StaleBookSetup)) else 400

    @app.errorhandler(sqlite3.Error)
    @app.errorhandler(OSError)
    def local_io_error(error):
        target = "数据库" if isinstance(error, sqlite3.Error) else "文件"
        return render_template("error.html", message=f"无法访问本地{target}。请检查权限、可用空间及路径后重试。"), 503

    @app.errorhandler(HTTPException)
    def http_error(error: HTTPException):
        messages = {400: "请求无效，请回到首页重试。", 403: "请求验证失败，请刷新页面后重试。",
                    404: "找不到此页面。", 413: "上传或表单超过大小限制。"}
        return render_template("error.html", message=messages.get(error.code, "此操作无法完成。")), error.code

    @app.get("/")
    def home():
        return render_template("home.html", books=configured_books().values(), allow_setup=allow_setup,
                               current_period=date.today().strftime("%Y-%m"))

    @app.get("/history")
    def history():
        book_filter = request.args.get("book", "")
        account, period = request.args.get("account", ""), request.args.get("period", "")
        books = configured_books()
        if book_filter and book_filter not in books:
            abort(404)
        selected = [books[book_filter]] if book_filter else list(books.values())
        records = [(book, row) for book in selected for row in store.recent_batches(book.id)
                   if (not account or row["account_id"] == account) and (not period or row["period"] == period)]
        records.sort(key=lambda item: item[1]["id"], reverse=True)
        return render_template("history.html", records=records, books=books.values(), book_filter=book_filter,
                               account=account, period=period)

    @app.get("/configuration")
    def configuration():
        return render_template("configuration.html", books=configured_books().values(), allow_setup=allow_setup)

    @app.get("/pending")
    def pending_accounting():
        books = configured_books()
        book_filter, period, group = (request.args.get(key, "") for key in ("book", "period", "group"))
        if book_filter and book_filter not in books:
            abort(404)
        if group and group not in PENDING_GROUPS:
            raise AccountingError("请选择有效的待确认类型。")
        if period:
            Period.from_text(period)
        page = _positive(request.args.get("page", "1"), "页码")
        page_size = 30
        items = store.pending_accounting(
            (book_filter,) if book_filter else tuple(books), period=period or None, group=group,
            limit=page_size + 1, offset=(page - 1) * page_size,
        )
        return render_template(
            "pending.html", books=books, book_filter=book_filter, period=period, group=group, groups=PENDING_GROUPS,
            items=items[:page_size], page=page, has_next=len(items) > page_size,
        )

    @app.get("/configuration/master-workbook")
    def download_setup_workbook():
        if not allow_setup:
            abort(404)
        return send_file(BytesIO(setup_workbook()), mimetype=MIME_XLSX, as_attachment=True,
                         download_name="账套主数据填写表.xlsx", max_age=0)

    @app.route("/configuration/new", methods=["GET", "POST"])
    def new_book():
        if not allow_setup:
            abort(404)
        if request.method == "GET":
            return render_template("book_setup.html", values={"word": "记", "currency_code": "RMB"})
        return submit_book_setup()

    def managed_book(book_id: str) -> BookSettings:
        book = get_book(book_id)
        if not book.managed:
            abort(404)
        return book

    @app.route("/books/<book_id>/batches/<int:batch_id>/recalculate/<int:row>", methods=["GET", "POST"])
    def recalculate_row(book_id: str, batch_id: int, row: int):
        book = managed_book(book_id)
        if request.method == "POST":
            if request.form.get("verified") != "yes":
                raise AccountingError("请核对试算分录，并确认将这条流水转入独立复核批次。")
            try:
                reviewed = recalculation_signer.loads(request.form.get("preview_token", ""), max_age=1800)
            except BadData as exc:
                raise AccountingError("试算结果无效或已过期，请重新试算。") from exc
            if (reviewed["book"], reviewed["batch"], reviewed["row"]) != (book.id, batch_id, row):
                raise AccountingError("试算结果不属于此来源行，请重新试算。")
            target, draft = store.apply_recalculation(
                book.id, batch_id, row, revision=reviewed["revision"], version=reviewed["version"],
            )
            flash("已生成重算凭证草稿，请在新批次中复核。原批次其他凭证不变。")
            return redirect(url_for("batch_page", book_id=book.id, batch_id=target, focus=draft,
                                    _anchor=f"draft-{draft}"), code=303)
        trial = store.preview_recalculation(book.id, batch_id, row)
        token = recalculation_signer.dumps({
            "book": book.id, "batch": batch_id, "row": row, "revision": trial.revision, "version": trial.ledger.version,
        }) if trial.voucher is not None else None
        return render_template("recalculate.html", book=book, trial=trial, token=token)

    @app.get("/configuration/books/<book_id>/master-workbook")
    def download_book_workbook(book_id: str):
        book = managed_book(book_id)
        config = store.book_setup_config(book.id, book.version)
        return send_file(BytesIO(setup_workbook(config)), mimetype=MIME_XLSX, as_attachment=True,
                         download_name=f"{book.id}-主数据-v{book.version}.xlsx", max_age=0)

    @app.route("/configuration/books/<book_id>/edit", methods=["GET", "POST"])
    def edit_book(book_id: str):
        book = managed_book(book_id)
        if request.method == "GET":
            return render_template(
                "book_setup.html", book=book, values={
                    "label": book.label, "word": book.word, "currency_code": book.currency_code,
                    "text_number": "yes" if book.text_number else "",
                },
            )
        return submit_book_setup(book)

    @app.route("/configuration/books/<book_id>/rules", methods=["GET", "POST"])
    def book_rules(book_id: str):
        book = managed_book(book_id)
        ledger, rules = parse_config(book.config)
        by_id = {rule.id: rule for rule in rules}

        def show(values, *, trial=None, token=None, error=None, status=200):
            batches = [row for row in store.recent_batches(book.id) if row["account_id"] in book.accounts]
            return render_template(
                "rule_setup.html", book=book, ledger=ledger, rules=rules, batches=batches, fields=MATCH_FIELDS,
                values=values, trial=trial, token=token, error=error,
            ), status

        if request.method == "GET":
            values = {"operation_type": "equals", "direction": "income", "batch_id": request.args.get("batch_id", ""),
                      "account_id": request.args.get("account_id", "")}
            if (rule_id := request.args.get("rule")) is not None:
                rule = by_id.get(rule_id)
                if rule is None or rule.counterpart is None:
                    abort(404)
                assignment = rule.counterpart
                values.update(rule_id=rule.id, editing=rule.id, account_id=rule.account_id or "", direction=rule.direction,
                              party=next(iter(rule.counterpart_names), ""), summary=rule.summary, subject=assignment.subject,
                              aux="\n".join(f"{aux.category}={aux.code}" for aux in assignment.auxiliary))
                if rule.conditions:
                    condition = rule.conditions[0]
                    values.update(field=condition.field, operation_type=condition.operation, text=condition.values[0])
            return show(values)

        operation = request.form.get("operation")
        if operation in ("publish", "disable"):
            if request.form.get("verified") != "yes":
                raise AccountingError("请核对变更范围并勾选确认。")
            if operation == "publish":
                try:
                    reviewed = rule_signer.loads(request.form.get("preview_token", ""), max_age=1800)
                except BadData as exc:
                    raise AccountingError("规则预览无效或已过期，请重新试跑。") from exc
                if reviewed["book"] != book.id:
                    raise AccountingError("规则预览不属于当前账套。")
                version, candidate = reviewed["version"], reviewed["rule"]
            else:
                version = _positive(request.form.get("version", ""), "账套配置版本")
                rule = by_id.get(request.form.get("rule_id", ""))
                if rule is None or not rule.enabled:
                    raise AccountingError("请选择当前已启用的规则。")
                candidate = next(raw for raw in book.config["rules"] if raw["id"] == rule.id)
                candidate = {**candidate, "enabled": False, "version": rule.version + 1}
            if version != book.version:
                raise StaleBookSetup("账套配置已更新，请基于最新版本重新操作。")
            config, template = store.book_setup(book.id, book.version)
            config["rules"] = [raw for raw in config["rules"] if raw["id"] != candidate["id"]] + [candidate]
            config["ledger"]["version"] = book.version + 1
            updated = _book_settings(config)
            store.publish_book_setup(book.id, tuple(item.account for item in updated.accounts.values()),
                                     config, template, previous_version=book.version)
            flash(f"规则「{candidate['id']}」已{'发布' if operation == 'publish' else '停用'}，仅影响新批次。")
            return redirect(url_for("book_rules", book_id=book.id), code=303)

        if operation != "preview":
            abort(400)
        values = {key: request.form.get(key, "").strip() for key in (
            "rule_id", "editing", "account_id", "direction", "party", "field", "operation_type",
            "text", "subject", "aux", "summary", "batch_id",
        )}
        try:
            if _positive(request.form.get("version", ""), "账套配置版本") != book.version:
                raise StaleBookSetup("账套配置已更新，请基于最新版本重新试跑。")
            previous = by_id.get(values["rule_id"])
            if (previous.id if previous is not None else "") != values["editing"]:
                raise AccountingError("规则标识已存在，或修改了正在编辑的规则标识。")
            if values["account_id"] not in book.accounts:
                raise AccountingError("请选择当前账套的银行账户。")
            _, assignment = _entry_fields()
            candidate_config = rule_config(book.id, values, assignment, previous.version + 1 if previous else 1)
            proposed_config = [raw for raw in book.config["rules"] if raw["id"] != values["rule_id"]] + [candidate_config]
            _, proposed = parse_config({"ledger": book.config["ledger"], "rules": proposed_config})
            validate_active_rules(ledger, proposed)
            trial = trial_rule(store, _positive(values["batch_id"], "试跑批次"), ledger, rules, proposed, proposed[-1])
            token = rule_signer.dumps({"book": book.id, "version": book.version, "rule": candidate_config}) if trial.can_publish else None
        except StaleBookSetup:
            raise
        except ValueError as exc:
            return show(values, error=str(exc), status=400)
        return show(values, trial=trial, token=token)

    @app.route("/configuration/books/<book_id>/accounts/<account_id>/format", methods=["GET", "POST"])
    def bank_format(book_id: str, account_id: str):
        book = managed_book(book_id)
        account = book.accounts.get(account_id)
        if account is None:
            abort(404)
        values = {"period": date.today().strftime("%Y-%m"), "amount_mode": "signed"}
        if request.method == "GET":
            config = store.book_setup_config(book.id, book.version)
            item = next(item for item in config["web"]["accounts"] if item["id"] == account_id)
            if "bank_profile" in item:
                profile = account.profiles[0]
                values.update({key: labels[0] for key, labels in profile.columns.items()})
                values.update(name=profile.name, amount_mode=profile.amount_mode, sheet_name=profile.sheet_name or "",
                              header_row=str(profile.header_row or ""), cny_only="yes" if profile.default_currency else "",
                              allow_serial_dates="yes" if profile.allow_serial_dates else "")
            return render_template("bank_setup.html", book=book, account=account, columns=COLUMNS, values=values)

        if request.form.get("operation") == "publish":
            if request.form.get("verified") != "yes":
                raise AccountingError("请核对预览笔数和金额，再勾选确认发布。")
            try:
                reviewed = preview_signer.loads(request.form.get("preview_token", ""), max_age=1800)
            except BadData as exc:
                raise AccountingError("银行格式预览无效或已过期，请重新上传样本预览。") from exc
            if (reviewed["book"], reviewed["account"]) != (book.id, account_id):
                raise AccountingError("预览不属于当前账套或银行账户，请重新预览。")
            if reviewed["version"] != book.version:
                raise StaleBookSetup("账套配置已更新，请重新打开银行格式配置并预览。")
            config, template = store.book_setup(book.id, book.version)
            item = next(item for item in config["web"]["accounts"] if item["id"] == account_id)
            item["bank_profile"] = reviewed["profile"]
            item.pop("profile_id", None)
            config["ledger"]["version"] = book.version + 1
            updated = _book_settings(config)
            store.publish_book_setup(updated.id, tuple(item.account for item in updated.accounts.values()),
                                     config, template, previous_version=book.version)
            flash(f"已绑定自定义银行格式「{item['bank_profile']['name']}」，仅供后续上传，账套配置 v{updated.version}。")
            return redirect(url_for("configuration"), code=303)

        if request.form.get("operation") != "preview":
            abort(400)
        values = {name: request.form.get(name, "").strip() for name in (
            *COLUMNS, "name", "amount_mode", "period", "sheet_name", "header_row", "allow_serial_dates", "cny_only",
        )}
        try:
            if _positive(request.form.get("version", ""), "账套配置版本") != book.version:
                raise StaleBookSetup("账套配置已更新，请重新打开银行格式配置并预览。")
            profile = custom_profile(values)
            sample = request.files.get("sample")
            if sample is None or not sample.filename:
                raise AccountingError("请选择与该账户对应的银行流水样本。")
            statement = parse_statement(
                sample.filename, sample.stream.read(LIMITS.file_bytes + 1), account=account.account,
                period=Period.from_text(values["period"]), posting_preference=account.posting_preference,
                profiles=(profile,),
            )
            preview = mapping_preview(statement)
            token = preview_signer.dumps({
                "book": book.id, "account": account_id, "version": book.version, "profile": asdict(profile),
            }) if preview.can_publish else None
        except StaleBookSetup:
            raise
        except ValueError as exc:
            return render_template("bank_setup.html", book=book, account=account, columns=COLUMNS, values=values,
                                   error=str(exc)), 400
        return render_template("bank_setup.html", book=book, account=account, columns=COLUMNS, values=values,
                               statement=statement, preview=preview, token=token)

    def submit_book_setup(existing: BookSettings | None = None):
        values = {name: request.form.get(name, "").strip() for name in (
            "book_id", "label", "word", "currency_code", "text_number", "verified",
        )}
        try:
            if existing is not None:
                values["book_id"] = existing.id
                if _positive(request.form.get("version", ""), "主数据版本") != existing.version:
                    raise StaleBookSetup("账套配置已更新，请刷新配置页并基于最新版本重新提交。")
            if values["verified"] != "yes":
                raise AccountingError("请核对主数据和凭证模板都对应同一个目标账套，并勾选财务核验。")
            if values["book_id"] in file_books:
                raise AccountingError("已经配置了这个账套标识，不能覆盖现有账套。")
            master, template = request.files.get("master"), request.files.get("template")
            if master is None or not master.filename:
                raise AccountingError("请选择账套主数据工作簿。")
            previous, previous_template = store.book_setup(existing.id, existing.version) if existing else ({}, None)
            if template is not None and template.filename:
                template_name, data = template.filename, template.stream.read(SETUP_LIMITS.file_bytes + 1)
            elif previous_template is not None:
                template_name, data = "existing.xlsx", previous_template
            else:
                raise AccountingError("请选择精斗云空白凭证模板。")
            existing_profiles = {item["id"]: item["bank_profile"] for item in previous.get("web", {}).get("accounts", [])
                                 if "bank_profile" in item}
            config = prepare_book_setup(
                values["book_id"], values["label"], values["word"], values["currency_code"],
                values["text_number"] == "yes", master.filename, master.stream.read(SETUP_LIMITS.file_bytes + 1),
                template_name, data, version=existing.version + 1 if existing is not None else 1,
                existing_profiles=existing_profiles,
            )
            if existing is not None:
                config["rules"] = existing.config["rules"]
            book = _book_settings(config)
            accounts = tuple(item.account for item in book.accounts.values())
            if existing is not None:
                store.publish_book_setup(book.id, accounts, config, data, previous_version=existing.version)
            else:
                store.create_book_setup(book.id, accounts, config, data)
        except StaleBookSetup:
            raise
        except ValueError as exc:
            return render_template("book_setup.html", book=existing, values=values, error=str(exc)), 400
        flash(f"已发布配置 v{book.version}，仅新批次使用。" if existing is not None else
              "账套已创建。请上传小批量流水，复核后再试导精斗云。")
        return redirect(url_for("configuration"), code=303)

    @app.post("/upload")
    def upload():
        book = get_book(request.form.get("book", ""))
        if (account := book.accounts.get(request.form.get("account", ""))) is None:
            raise AccountingError("请选择该账套下有效的银行账户。")
        file = request.files.get("statement")
        if file is None or not file.filename:
            raise AccountingError("请选择要上传的银行流水文件。")
        statement = parse_statement(
            file.filename, file.stream.read(LIMITS.file_bytes + 1), account=account.account,
            period=Period.from_text(request.form.get("period", "")), sheet=request.form.get("sheet", "").strip() or None,
            posting_preference=account.posting_preference, confirmed_binding=request.form.get("binding") == "yes",
            profiles=account.profiles,
        )
        saved = store.save(statement, expected_setup=book.version if book.managed else None)
        store.prepare(saved.id, book.config)
        if saved.reused:
            flash("文件已处理过，已打开原批次。")
        return redirect(url_for("batch_page", book_id=book.id, batch_id=saved.id), code=303)

    @app.get("/books/<book_id>/batches/<int:batch_id>")
    def batch_page(book_id: str, batch_id: int):
        book = get_book(book_id)
        store.prepare(batch_id, book.config)
        batch, ledger, drafts, exports = store.review_data(batch_id, book.id)
        if book.managed and book.version != ledger.version:
            book = _book_settings(store.book_setup_config(book.id, ledger.version))
        summary = store.summary(batch_id)
        exclusions = review_exclusions(batch)
        skipped = {item["row_number"] for item in exclusions if item.get("skipped")}

        def is_skipped(item):
            return any(source.row_number in skipped for source in item.voucher.sources)

        active = [item for item in drafts if item.export_id is None and not is_skipped(item)]
        ready = [item for item in active if item.original_export_id is None and item.voucher.review_digest is not None]
        pending = sum(item.voucher.review_digest is None for item in active)
        query, status, day = request.args.get("q", "").strip().casefold(), request.args.get("status", ""), request.args.get("date", "")
        filtered = []
        for item in drafts:
            tag = ("exported" if item.export_id is not None else "skipped" if is_skipped(item)
                   else "reviewed" if item.voucher.review_digest else "pending")
            if status and tag != status or day and item.voucher.sources[0].transaction.voucher_date.isoformat() != day:
                continue
            texts = [entry.summary for entry in item.voucher.entries]
            for source in item.voucher.sources:
                texts.extend((source.transaction.counterparty_name or "", source.transaction.bank_reference or "", source.transaction.summary or ""))
            if query and not any(query in text.casefold() for text in texts):
                continue
            filtered.append((item, tag))
        page = max(1, min(request.args.get("page", 1, type=int), max(1, (len(filtered) + 19) // 20)))
        if (focus := request.args.get("focus", type=int)) is not None:
            page = next((index // 20 + 1 for index, (item, _) in enumerate(filtered) if item.id == focus), page)
        visible = [(draft, tag, validate_voucher(draft.voucher, ledger)) for draft, tag in filtered[(page - 1) * 20:page * 20]]
        retry_requests = {draft.id: secrets.token_urlsafe(24) for draft, _, _ in visible if draft.original_export_id}
        return render_template("batch.html", book=book, batch=batch, ledger=ledger, summary=summary, exports=exports,
                               exclusions=exclusions, skipped=skipped, skip_reasons=SKIP_REASONS,
                               exception_row=request.args.get("exception_row", type=int),
                               manual_rows={item["row_number"] for item in exclusions if can_resolve_manually(item)},
                               drafts=visible, retry_requests=retry_requests, ready=len(ready), pending=pending,
                               ready_entries=sum(len(item.voucher.entries) for item in ready),
                               excluded_rows=sorted({item["row_number"] for item in exclusions} |
                                                    {s.row_number for item in active if item.original_export_id is None and not item.voucher.review_digest
                                                     for s in item.voucher.sources}),
                               total=len(drafts), request_id=secrets.token_urlsafe(24), page=page,
                               has_next=page * 20 < len(filtered), query=query, status=status, day=day)

    @app.post("/books/<book_id>/batches/<int:batch_id>/drafts/<int:draft_id>")
    def change_draft(book_id: str, batch_id: int, draft_id: int):
        book = get_book(book_id)
        token = request.form.get("token", "")
        operation = request.form.get("operation")
        if operation == "confirm":
            store.change_draft(book.id, batch_id, draft_id, token)
            flash("凭证已复核。")
        elif operation == "edit":
            index = _positive(request.form.get("line", ""), "分录序号") - 1
            store.change_draft(book.id, batch_id, draft_id, token, edits={index: _entry_fields()})
            flash("分录已保存，需重新复核。")
        else:
            abort(400)
        return redirect(url_for("batch_page", book_id=book.id, batch_id=batch_id, focus=draft_id,
                                _anchor=f"draft-{draft_id}"), code=303)

    @app.route("/books/<book_id>/batches/<int:batch_id>/manual/<int:row>", methods=["GET", "POST"])
    def manual_draft(book_id: str, batch_id: int, row: int):
        book = get_book(book_id)
        error = None
        if request.method != "POST":
            store.prepare(batch_id, book.config)
        else:
            try:
                summary, assignment = _entry_fields()
                draft_id = store.create_manual(book.id, batch_id, _revision(), row, assignment, summary,
                                               request.form.get("reason", ""))
            except AccountingError as exc:
                error = str(exc)
            else:
                flash("已创建手工凭证草稿，请核对分录并单独确认。")
                return redirect(url_for("batch_page", book_id=book.id, batch_id=batch_id, focus=draft_id,
                                        _anchor=f"draft-{draft_id}"), code=303)
        revision, ledger, source, exclusion = store.manual_details(book.id, batch_id, row)
        bank_subject = ledger.bank_subjects[source.account_id]
        subjects = [subject for code, subject in sorted(ledger.subjects.items())
                    if code != bank_subject and not subject.quantity_required]
        return render_template("manual.html", book=book, batch_id=batch_id, source=source, exclusion=exclusion,
                               revision=revision, bank_subject=ledger.subjects[bank_subject], subjects=subjects, error=error,
                               submitted=request.form), 400 if error else 200

    @app.post("/books/<book_id>/batches/<int:batch_id>/skip/<int:row>")
    def skip_row(book_id: str, batch_id: int, row: int):
        book = get_book(book_id)
        undo = request.form.get("operation") == "undo"
        store.set_skip(book.id, batch_id, _revision(), row, None if undo else request.form.get("reason", ""),
                       request.form.get("note", ""))
        flash("已撤回跳过，请重新复核。" if undo else f"已跳过原流水第 {row} 行。")
        return redirect(url_for("batch_page", book_id=book.id, batch_id=batch_id), code=303)

    @app.post("/books/<book_id>/batches/<int:batch_id>/regroup")
    def regroup(book_id: str, batch_id: int):
        book = get_book(book_id)
        operation = request.form.get("operation")
        if operation not in ("split", "merge"):
            abort(400)
        ids = tuple(_positive(text, "凭证编号") for text in re.split(r"[\s,，、]+", request.form.get("ids", "")) if text)
        created = store.regroup(book.id, batch_id, _revision(), ids, split=operation == "split")
        flash("凭证已拆分，请重新复核。" if operation == "split" else "凭证已合并，请重新复核。")
        return redirect(url_for("batch_page", book_id=book.id, batch_id=batch_id, focus=created[0]), code=303)

    @app.post("/books/<book_id>/batches/<int:batch_id>/export")
    def export_batch(book_id: str, batch_id: int):
        book = get_book(book_id)
        batch, ledger, _, _ = store.review_data(batch_id, book.id)
        if not book.can_export or not ledger.finance_verified:
            raise AccountingError("当前账套或本批主数据未核验，不能生成实际导入文件。")
        start = _positive(request.form.get("first_number", ""), "起始凭证号")
        used = frozenset(_positive(n, "既用凭证号") for n in re.split(r"[\s,，、]+", request.form.get("used", "")) if n)
        request_id, retry = request.form.get("request_id", ""), request.form.get("retry_draft_id", "")
        if not re.fullmatch(r"[A-Za-z0-9_-]{32}", request_id):
            abort(400)
        if book.managed:
            setup, template = store.book_setup(book.id, ledger.version)
            book = _book_settings(setup)
        elif book.template is not None:
            with book.template.open("rb") as file:
                template = file.read(LIMITS.file_bytes + 1)
        else:
            raise AccountingError("缺少已核验的目标账套模板。")
        options = ExportOptions(Period.from_text(batch["period"]), book.word, start, book.currency_code, book.text_number,
                                used, book.template_hash)
        export_id = store.create_export(book.id, batch_id, _revision(), request_id, template, options,
                                        allow_partial=request.form.get("partial") == "yes",
                                        retry_draft_id=_positive(retry, "修订凭证") if retry else None)
        return redirect(url_for("export_page", book_id=book.id, export_id=export_id), code=303)

    @app.get("/books/<book_id>/exports/<int:export_id>")
    def export_page(book_id: str, export_id: int):
        book = get_book(book_id)
        export, items = store.export_data(book.id, export_id)
        return render_template("export.html", book=book, export=export, items=items,
                               has_retries=any(item["retry_draft_id"] is not None for item, _ in items),
                               exclusions=json.loads(export["exclusions_json"]))

    @app.get("/books/<book_id>/exports/<int:export_id>/download")
    def download(book_id: str, export_id: int):
        book = get_book(book_id)
        reconciliation = request.args.get("file") == "reconciliation"
        label = "核对底稿（勿导入）" if reconciliation else "精斗云导入"
        return send_file(BytesIO(store.export_file(book.id, export_id, reconciliation=reconciliation)),
                         mimetype=MIME_XLSX, as_attachment=True, download_name=f"{label}-{export_id}.xlsx", max_age=0)

    @app.post("/books/<book_id>/exports/<int:export_id>/numbers/<int:number>")
    def import_result(book_id: str, export_id: int, number: int):
        book = get_book(book_id)
        store.record_import(book.id, export_id, number, request.form.get("status", ""), request.form.get("note", ""))
        flash("已登记精斗云导入结果。")
        return redirect(url_for("export_page", book_id=book.id, export_id=export_id), code=303)

    @app.post("/books/<book_id>/exports/<int:export_id>/numbers/<int:number>/retry")
    def start_retry(book_id: str, export_id: int, number: int):
        book = get_book(book_id)
        if request.form.get("confirm") != "yes":
            raise AccountingError("请确认已核实精斗云未入账，并停止使用该张原导出凭证。")
        batch_id, draft_id = store.start_retry(book.id, export_id, number)
        flash("已打开单张修订凭证。请修改并重新复核，然后从该凭证卡片单独导出。")
        return redirect(url_for("batch_page", book_id=book.id, batch_id=batch_id, focus=draft_id,
                                _anchor=f"draft-{draft_id}"), code=303)

    return app


def serve(database: Path, configs: list[Path], port: int) -> None:
    from waitress import serve as waitress_serve

    if not 1 <= port <= 65535:
        raise ValueError("端口号必须在 1 至 65535 之间。")
    app = create_app(database, configs)
    print(f"请在浏览器打开 http://127.0.0.1:{port} （关闭终端可停止）", flush=True)
    waitress_serve(app, host="127.0.0.1", port=port, max_request_body_size=UPLOAD_LIMIT)
