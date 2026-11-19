#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从「原版表格.docx」生成 docxtpl 内部模板与结构清单。

运行：
    python make_templates.py

输入：原版表格.docx（与脚本同目录）
输出：_内部模板.docx、_模板结构.json
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import docx  # noqa: F401
    import openpyxl  # noqa: F401
except ImportError:
    print("缺少运行依赖（python-docx / openpyxl）。")
    print("请先执行：python -m pip install -r requirements.txt")
    raise SystemExit(1)

from settlement.config import load_config  # noqa: E402
from settlement.constants import (  # noqa: E402
    CONFIG_FILE,
    INTERNAL_TEMPLATE,
    MANIFEST_FILE,
    PROJECT_TITLE,
    SOURCE_DOCX,
    VERSION,
)
from settlement.errors import SettlementError  # noqa: E402
from settlement.template import build_template, check_template  # noqa: E402
from settlement.utils import (  # noqa: E402
    banner,
    error,
    hint,
    info,
    ok,
    resolve_path,
    warn,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="make_templates.py",
        description="把原版表格.docx 转换为 docxtpl 内部模板。",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--root", default=str(ROOT), help="工作目录，默认脚本所在目录")
    parser.add_argument("--source", default=SOURCE_DOCX, help="原版表格.docx")
    parser.add_argument("--output", default=INTERNAL_TEMPLATE, help="输出的内部模板")
    parser.add_argument("--manifest", default=MANIFEST_FILE, help="输出的模板结构清单")
    parser.add_argument(
        "--table-index",
        type=int,
        default=None,
        help="指定第几个表格（1-based）；默认按配置或自动选择",
    )
    parser.add_argument(
        "--header-rows",
        type=int,
        default=None,
        help="表头占几行；默认读取规则配置，缺省为 1",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="只检查模板契约，不生成内部模板",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = Path(args.root).expanduser().resolve()
    source = resolve_path(root, args.source, SOURCE_DOCX)
    target = resolve_path(root, args.output, INTERNAL_TEMPLATE)
    manifest = resolve_path(root, args.manifest, MANIFEST_FILE)

    banner(PROJECT_TITLE, "生成 Word 内部模板")
    info(f"工作目录：{root}")
    info(f"原版表格：{source.name}")

    if not source.exists():
        error(f"未找到「{source.name}」：{source}")
        hint("请把原始 Word 表格命名为「原版表格.docx」放到本目录。")
        return 1

    header_rows = args.header_rows
    table_index = args.table_index
    if header_rows is None or table_index is None:
        config_path = root / CONFIG_FILE
        if config_path.exists():
            try:
                config = load_config(config_path)
                for message in config.warnings:
                    warn(message)
                if header_rows is None:
                    header_rows = config.base.header_rows
                if table_index is None:
                    table_index = config.base.table_index
            except SettlementError as exc:
                warn(f"读取规则配置失败，使用默认模板参数：{exc}")
        else:
            warn(f"未找到「{config_path.name}」，使用默认表头行数 1、自动选择表格。")

    header_rows = 1 if header_rows is None else header_rows
    if header_rows < 1:
        error("「表头行数」至少为 1。")
        return 1

    if args.check:
        try:
            report = check_template(
                source,
                table_index=table_index,
                header_rows=header_rows,
            )
        except SettlementError as exc:
            error(str(exc))
            return 1
        except OSError as exc:
            error(f"读取模板失败：{exc}")
            return 1

        for line in report.summary_lines():
            info(line)
        for message in report.warnings:
            warn(message)
        if report.ok:
            ok("模板检查通过：可以生成内部模板。")
            return 0
        for message in report.errors:
            error(message)
        hint("修正上述问题后，重新运行 python make_templates.py --check。")
        return 1

    try:
        result = build_template(
            source,
            target,
            manifest,
            table_index=table_index,
            header_rows=header_rows,
        )
    except SettlementError as exc:
        error(str(exc))
        return 1
    except OSError as exc:
        error(f"写入模板失败：{exc}")
        hint("请关闭正在占用「原版表格.docx」或「_内部模板.docx」的 Word / WPS 窗口。")
        return 1

    for message in result.warnings:
        warn(message)

    ok(f"已生成内部模板：{target.name}")
    ok(f"已生成模板结构：{manifest.name}")
    info(
        f"主表格：第 {result.manifest['table_index']} 个；"
        f"表头 {result.manifest['header_rows']} 行；"
        f"数据列 {len(result.manifest['headers'])} 列"
    )
    info("列名：" + " | ".join(result.manifest["headers"]))
    if result.manifest.get("total_slots"):
        info(
            "表尾合计列："
            + "、".join(item["header"] for item in result.manifest["total_slots"])
        )
    fonts = result.manifest.get("fonts") or {}
    info(
        "字体识别：表头 "
        + _font_profile_text(fonts.get("header") or {})
        + "；数据 "
        + _font_profile_text(fonts.get("data") or {})
    )
    hint("接下来：确认「源数据.xlsx」已填好，然后运行 make_report.py 或双击 一键运行.bat。")
    return 0


def _font_profile_text(profile: dict) -> str:
    if not profile:
        return "未显式设置（跟随 Word 默认）"
    parts: list[str] = []
    if profile.get("east_asia"):
        parts.append(f"中文 {profile['east_asia']}")
    if profile.get("ascii"):
        parts.append(f"西文 {profile['ascii']}")
    if profile.get("size_pt"):
        parts.append(f"{float(profile['size_pt']):g} 磅")
    return " / ".join(parts) or "未显式设置（跟随 Word 默认）"


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已取消。")
        raise SystemExit(130)
