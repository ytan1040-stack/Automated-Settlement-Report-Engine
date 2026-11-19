#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""用 PyInstaller 构建 Windows 便携版（one-folder）。

首次使用：
    python -m pip install -r requirements.txt pyinstaller
    python build_exe.py

产物：
    dist/一键运行/一键运行.exe

把 原版表格.docx、源数据.xlsx、规则配置.xlsx 放到 exe 同目录，双击 exe 即可。
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="构建 Windows 便携版 exe。")
    parser.add_argument("--name", default="一键运行", help="exe 名称，默认 一键运行")
    parser.add_argument("--no-clean", action="store_true", help="保留 PyInstaller 缓存")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if importlib.util.find_spec("PyInstaller") is None:
        print("未安装 PyInstaller。请先执行：")
        print("  python -m pip install pyinstaller")
        return 1

    spec_dir = ROOT / "build" / "pyinstaller"
    work_dir = spec_dir / "work"
    dist_dir = ROOT / "dist"
    spec_dir.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--console",
        "--name",
        args.name,
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        "--collect-all",
        "docxtpl",
        "--collect-all",
        "docx",
        "--hidden-import",
        "lxml.etree",
        "--hidden-import",
        "lxml._elementpath",
        str(ROOT / "portable_entry.py"),
    ]
    if not args.no_clean:
        command.insert(4, "--clean")

    print("执行：")
    print("  " + " ".join(command))
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        print("构建失败：PyInstaller 返回非零退出码。")
        return result.returncode

    target = dist_dir / args.name
    readme = ROOT / "README.md"
    if readme.exists():
        shutil.copy2(readme, target / "README.md")
    license_file = ROOT / "LICENSE"
    if license_file.exists():
        shutil.copy2(license_file, target / "LICENSE")

    print("")
    print("构建完成：")
    print(f"  {target / (args.name + '.exe')}")
    print("")
    print("使用方法：")
    print(f"  1. 把 原版表格.docx、源数据.xlsx、规则配置.xlsx 放到 {target}")
    print("  2. 双击 一键运行.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
