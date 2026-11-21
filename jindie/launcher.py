from __future__ import annotations

import argparse
import sqlite3
import sys
import webbrowser
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from waitress import create_server

from .storage import default_database
from .web import UPLOAD_LIMIT, create_app


class LaunchError(ValueError):
    """A recoverable problem opening the local desktop workspace."""


@contextmanager
def _workspace_lock(database: Path) -> Iterator[None]:
    import msvcrt

    database.parent.mkdir(parents=True, exist_ok=True)
    with database.with_suffix(".lock").open("a+b") as lock:
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise LaunchError("工作区可能已启动；请查看原浏览器窗口，或关闭原启动窗口后重试。") from exc
        try:
            yield
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


def _settings(bundle_root: Path, data_dir: Path, demo: bool) -> tuple[Path, list[Path]]:
    if demo:
        config = bundle_root / "demo" / "voucher-config.json"
        if not config.is_file():
            raise LaunchError("演示配置缺失，请重新解压整个启动包。")
        return data_dir / "demo" / "statements.sqlite3", [config]
    books = data_dir / "books"
    books.mkdir(parents=True, exist_ok=True)
    configs = sorted(item for item in books.glob("*.json") if item.is_file())
    return data_dir / "statements.sqlite3", configs


def main() -> int:
    parser = argparse.ArgumentParser(description="启动银行流水凭证工作台")
    parser.add_argument("--bundle-root", required=True, type=Path)
    parser.add_argument("--demo", action="store_true")
    parser.add_argument("--no-browser", action="store_true", help="仅打印网址，供离线验证")
    args = parser.parse_args()
    data_dir = default_database().parent
    try:
        database, configs = _settings(args.bundle_root.resolve(), data_dir, args.demo)
        with _workspace_lock(database):
            app = create_app(database, configs, allow_setup=not args.demo)
            server = create_server(app, host="127.0.0.1", port=0, max_request_body_size=UPLOAD_LIMIT)
            url = f"http://127.0.0.1:{server.effective_port}/"
            print(url, flush=True)
            print("关闭此窗口可停止服务；已保存的批次不会丢失。", flush=True)
            try:
                if not args.no_browser:
                    try:
                        opened = webbrowser.open(url)
                    except webbrowser.Error:
                        opened = False
                    if not opened:
                        print("未能自动打开浏览器，请将上方地址复制到 Chrome 或 Edge。", flush=True)
                server.run()
            finally:
                server.close()
    except LaunchError as exc:
        print(f"启动未完成：{exc}", file=sys.stderr)
        return 2
    except (OSError, sqlite3.Error, ValueError):
        print(f"启动未完成：请检查 {data_dir} 中的账套配置、文件权限与数据库。"
              "如果迁移失败，请保留原库和 .bak 备份并联系维护人员。", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
