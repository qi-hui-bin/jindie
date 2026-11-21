import argparse
import getpass
import json
import sqlite3
import sys
from pathlib import Path

from .bank_import import parse_statement
from .bank_profiles import BUILTIN_PROFILES, BankProfile
from .excel_reader import ImportLimits
from .models import BookAccount, Period
from .storage import StatementStore, default_database
from .voucher_preview import create_preview


def main() -> int:
    parser = argparse.ArgumentParser(description="银行流水解析与批次检查（开发入口）")
    parser.add_argument("--db", type=Path, default=default_database(), help="本地 SQLite 路径")
    actions = parser.add_subparsers(dest="action", required=True)
    importer = actions.add_parser("import", help="解析并保存一份银行流水")
    importer.add_argument("file", type=Path)
    importer.add_argument("--book", required=True)
    importer.add_argument("--account-id", required=True)
    importer.add_argument("--account-number", help="仅建议给合成样本使用；省略时安全输入完整账号")
    importer.add_argument("--period", required=True, metavar="YYYY-MM")
    formats = importer.add_mutually_exclusive_group()
    formats.add_argument("--profile", choices=[profile.id for profile in BUILTIN_PROFILES])
    formats.add_argument("--profile-json", type=Path, help="自定义银行字段映射")
    importer.add_argument("--sheet")
    importer.add_argument("--posting-preference", choices=("transaction", "posting"))
    importer.add_argument("--confirm-binding", action="store_true", help="已核实缺失或遮蔽的本方账号")
    inspector = actions.add_parser("inspect", help="查看已保存批次的汇总和逐行问题")
    inspector.add_argument("batch_id", type=int)
    preview = actions.add_parser("preview", help="生成凭证试算工作簿（不可导入）")
    preview.add_argument("batch_id", type=int)
    preview.add_argument("--config", required=True, type=Path)
    preview.add_argument("--template", required=True, type=Path)
    preview.add_argument("--output", required=True, type=Path)
    preview.add_argument("--word", default="记")
    preview.add_argument("--first-number", type=int, required=True)
    preview.add_argument("--used-number", type=int, action="append", default=[])
    preview.add_argument("--text-number", action="store_true")
    preview.add_argument("--confirmed-skips", type=Path, help="财务已核实的同批流水号与原因")
    preview.add_argument("--include-partial", action="store_true", help="允许仅预览通过校验的凭证，排除项继续单列")
    server = actions.add_parser("serve", help="启动本地浏览器界面")
    server.add_argument("--config", default=[], type=Path, action="append", help="已有的本地账套 JSON，可重复提供")
    server.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    store = StatementStore(args.db)
    try:
        if args.action == "serve":
            from .web import serve

            serve(args.db, args.config, args.port)
            return 0
        if args.action == "inspect":
            report = store.summary(args.batch_id)
        elif args.action == "preview":
            report = create_preview(
                store, args.batch_id, args.config, args.template, args.output, args.word, args.first_number,
                skip_file=args.confirmed_skips, include_partial=args.include_partial, text_number=args.text_number,
                used_numbers=frozenset(args.used_number),
            )
        else:
            limits = ImportLimits()
            profiles = BUILTIN_PROFILES
            if args.profile:
                profiles = tuple(profile for profile in profiles if profile.id == args.profile)
            if args.profile_json:
                try:
                    config = json.loads(args.profile_json.read_text(encoding="utf-8-sig"))
                    profiles = (BankProfile(**config),)
                except (ValueError, TypeError, AttributeError) as exc:
                    raise ValueError("银行映射 JSON 格式无效，请检查字段、别名和版本。") from exc
            account = BookAccount(
                args.book, args.account_id, args.account_number or getpass.getpass("完整银行账号（输入不回显）：")
            )
            with args.file.open("rb") as source:
                data = source.read(limits.file_bytes + 1)
            result = parse_statement(
                args.file.name, data, account=account, period=Period.from_text(args.period), profiles=profiles,
                sheet=args.sheet, posting_preference=args.posting_preference, confirmed_binding=args.confirm_binding,
                limits=limits,
            )
            saved = store.save(result)
            report = store.summary(saved.id)
            report["reused"] = saved.reused
    except ValueError as exc:
        print(f"处理失败：{exc}", file=sys.stderr)
        return 2
    except EOFError:
        print("无法读取完整账号，请在交互式终端运行。", file=sys.stderr)
        return 2
    except FileExistsError:
        print("输出文件已存在，不会覆盖；请选择新文件名。", file=sys.stderr)
        return 2
    except OSError:
        print("无法读取输入文件或写入数据库，请检查路径与权限。", file=sys.stderr)
        return 2
    except sqlite3.Error:
        print("数据库操作失败，请检查数据库是否正被占用、是否可写或需要恢复。", file=sys.stderr)
        return 2
    print(json.dumps(report, ensure_ascii=True, indent=2))
    if args.action == "preview" and not report["workbook_written"]:
        print("没有生成预览：请处理排除项，或在有有效凭证时显式允许部分预览。", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
