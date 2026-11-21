"""Build a portable Windows distribution using the official CPython embeddable archive."""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

from tests.voucher_fixtures import write_example

ROOT = Path(__file__).resolve().parents[1]


def build(runtime_zip: Path, output: Path, runtime_sha256: str | None = None) -> None:
    if sys.platform != "win32":
        raise ValueError("请在 Windows 上构建并验证启动包。")
    if output.suffix.lower() != ".zip" or output.exists():
        raise ValueError("输出必须是一个尚不存在的 .zip 路径。")
    digest = sha256()
    with runtime_zip.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    signature = digest.hexdigest()
    if runtime_sha256 is not None and signature != runtime_sha256.lower():
        raise ValueError("Python 运行时 ZIP 的 SHA-256 与预期值不一致。")
    print(f"CPython runtime SHA-256: {signature}", flush=True)

    with TemporaryDirectory(prefix="jindie-windows-") as work:
        bundle = Path(work) / "Jindie"
        runtime = bundle / "runtime"
        runtime.mkdir(parents=True)
        with ZipFile(runtime_zip) as source:
            members = source.namelist()
            if not {"python.exe", "python312.dll", "python312._pth", "python312.zip"} <= set(members):
                raise ValueError("请提供 python.org 的 Windows 3.12 x64 embeddable ZIP。")
            if any(Path(name).name != name or name in (".", "..") or ":" in name or "/" in name or "\\" in name
                   for name in members):
                raise ValueError("Python 运行时压缩包包含异常路径。")
            source.extractall(runtime)
        python = runtime / "python.exe"
        subprocess.run(
            [str(python), "-c", "import struct, sys; assert sys.version_info[:2] == (3, 12) and struct.calcsize('P') == 8"],
            check=True,
        )
        target = runtime / "Lib" / "site-packages"
        subprocess.run([
            sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-deps", "--no-compile", "--only-binary=:all:",
            "--target", str(target), "--platform", "win_amd64", "--python-version", "3.12", "--implementation", "cp",
            "--abi", "cp312", "-r", str(ROOT / "requirements.lock"),
        ], check=True)
        (runtime / "python312._pth").write_text("python312.zip\n.\nLib\\site-packages\n..\\app\n", encoding="ascii")
        shutil.copytree(ROOT / "jindie", bundle / "app" / "jindie", ignore=shutil.ignore_patterns("__pycache__"))
        for script in ("Start-Jindie.cmd", "Start-Demo.cmd"):
            shutil.copy2(ROOT / "windows" / script, bundle / script)
        for document in ("README.md", "CURRENT_SCOPE.md", "PLAN.md"):
            shutil.copy2(ROOT / document, bundle / document)
        write_example(bundle / "demo")
        (bundle / "demo" / "confirmed-skips.json").unlink()
        subprocess.run([str(python), "-c", (
            "from importlib.resources import files; from jindie.web import create_app; import xlrd, openpyxl; "
            "assert files('jindie').joinpath('templates', 'base.html').is_file()"
        )], check=True)
        staged = Path(work) / "distribution.zip"
        with ZipFile(staged, "w", ZIP_DEFLATED) as archive:
            for item in sorted(bundle.rglob("*")):
                if item.is_file() and "__pycache__" not in item.parts:
                    archive.write(item, item.relative_to(bundle.parent))
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("xb") as destination, staged.open("rb") as source:
            shutil.copyfileobj(source, destination)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runtime-zip", type=Path, required=True, help="Official CPython 3.12 x64 embeddable archive")
    parser.add_argument("--runtime-sha256", help="Expected SHA-256 of the official runtime archive")
    parser.add_argument("--output", type=Path, default=ROOT / "dist" / "Jindie-Windows-x64.zip")
    args = parser.parse_args()
    try:
        build(args.runtime_zip, args.output, args.runtime_sha256)
    except (OSError, ValueError, BadZipFile, subprocess.CalledProcessError) as exc:
        parser.exit(2, f"无法生成 Windows 启动包：{exc}\n")
    print(args.output.resolve())


if __name__ == "__main__":
    main()
