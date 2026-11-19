#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""初始化规则配置与源数据模板。

双击或运行：
    python init_config.py

生成：
    规则配置.xlsx      三个 Sheet：基础配置 / 动态计费项 / 字段映射
    源数据模板.xlsx    标准列名 + 两行示例数据
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import openpyxl  # noqa: F401
except ImportError:
    print("缺少运行依赖 openpyxl。")
    print("请先执行：python -m pip install -r requirements.txt")
    raise SystemExit(1)

from settlement.config_writer import create_config, create_data_template  # noqa: E402
from settlement.constants import (  # noqa: E402
    CONFIG_FILE,
    DATA_FILE,
    DATA_TEMPLATE_FILE,
    PROJECT_TITLE,
    SOURCE_DOCX,
)
from settlement.errors import SettlementError  # noqa: E402
from settlement.utils import banner, error, hint, info, ok  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="生成规则配置.xlsx 与源数据模板.xlsx",
    )
    parser.add_argument(
        "--root",
        default=str(ROOT),
        help="工作目录，默认是脚本所在目录",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="覆盖已存在的规则配置.xlsx / 源数据模板.xlsx（会丢失旧配置）",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = Path(args.root).expanduser().resolve()

    banner(PROJECT_TITLE, "初始化配置")
    info(f"工作目录：{root}")
    if not root.exists():
        error(f"工作目录不存在：{root}")
        return 1

    try:
        if args.force:
            hint("--force 已开启：将覆盖已有的配置与模板文件")
        config_path = create_config(root / CONFIG_FILE, overwrite=args.force)
        template_path = create_data_template(
            root / DATA_TEMPLATE_FILE,
            overwrite=args.force,
        )
    except SettlementError as exc:
        error(str(exc))
        hint("请关闭正在占用文件的 Excel / WPS 窗口后重试。")
        return 1
    except OSError as exc:
        error(f"写入文件失败：{exc}")
        return 1

    ok("初始化完成")
    info(f"配置文件：{config_path.name}")
    info(f"数据模板：{template_path.name}")

    if not (root / DATA_FILE).exists():
        hint("下一步：按「源数据模板.xlsx」填好数据，另存为「源数据.xlsx」。")
    if not (root / SOURCE_DOCX).exists():
        hint("下一步：把原始 Word 表格命名为「原版表格.docx」放到本目录。")
    hint("最后双击「一键运行.bat」，自动生成结算报表。")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已取消。")
        raise SystemExit(130)
