# -*- coding: utf-8 -*-
"""命令行入口：读取配置与数据，计算并渲染结算报表。"""

from __future__ import annotations

import argparse
import os
import traceback
from pathlib import Path

from .constants import (
    CONFIG_FILE,
    DATA_FILE,
    INTERNAL_TEMPLATE,
    MANIFEST_FILE,
    SOURCE_DOCX,
    VERSION,
)
from .errors import FileLockedError, SettlementError
from .pipeline import RunOptions, run
from .utils import error, hint, info, ok, warn


def build_parser(default_root: Path | None = None) -> argparse.ArgumentParser:
    root = default_root or Path.cwd()
    parser = argparse.ArgumentParser(
        prog="make_report.py",
        description="读取源数据，按规则计算，并用 Word 模板生成结算报表。",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--root", default=str(root), help="工作目录，默认当前目录")
    parser.add_argument("--config", default=CONFIG_FILE, help="规则配置文件")
    parser.add_argument("--data", default=DATA_FILE, help="源数据文件")
    parser.add_argument("--source", default=SOURCE_DOCX, help="原版表格.docx")
    parser.add_argument("--template", default=INTERNAL_TEMPLATE, help="内部模板文件")
    parser.add_argument("--manifest", default=MANIFEST_FILE, help="模板结构文件")
    parser.add_argument("--output", default=None, help="输出文件名（默认 结算报表_YYYYMMDD.docx）")
    parser.add_argument("--sheet", default=None, help="源数据工作表名（默认第一个）")
    parser.add_argument(
        "--rebuild-template",
        action="store_true",
        help="强制根据原版表格重新生成内部模板",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只读取、计算和校验，不写出 Word 文件",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="生成 Word 前先写出「数据预览.xlsx」（数据表、合计、页数），不生成 Word",
    )
    parser.add_argument(
        "--preview-rows",
        type=int,
        default=10,
        help="预览时控制台显示的行数，默认 10；完整数据在 数据预览.xlsx",
    )
    parser.add_argument(
        "--compare",
        action="store_true",
        help="与上一期成功运行做差异对比，写出「差异报告.xlsx」",
    )
    parser.add_argument(
        "--compare-run",
        default=None,
        help="指定要对比的历史运行 ID（runs 目录下的名字），隐含 --compare",
    )
    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="跳过数据校验与数据检查报告",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="覆盖输出文件前不备份旧文件",
    )
    return parser


def main(
    argv: list[str] | None = None,
    default_root: Path | None = None,
) -> int:
    args = build_parser(default_root).parse_args(argv)
    root = Path(args.root).expanduser().resolve()

    options = RunOptions(
        root=root,
        config_path=args.config,
        data_path=args.data,
        source_docx=args.source,
        template_path=args.template,
        manifest_path=args.manifest,
        output_path=args.output,
        rebuild_template=args.rebuild_template,
        sheet_name=args.sheet,
        dry_run=args.dry_run or args.preview,
        validate=not args.no_validate,
        backup=not args.no_backup,
        preview=args.preview,
        preview_rows=max(0, args.preview_rows),
        compare=args.compare or bool(args.compare_run),
        compare_run=args.compare_run,
    )

    try:
        result = run(options)
    except FileLockedError as exc:
        error(str(exc))
        hint("请关闭正在占用文件的 Word / WPS / Excel 窗口，然后重新运行。")
        return 2
    except SettlementError as exc:
        error(str(exc))
        return 1
    except KeyboardInterrupt:
        warn("已取消。")
        return 130
    except Exception as exc:  # noqa: BLE001 - 顶层兜底，避免双击窗口直接闪退
        error(f"发生未预期的错误：{exc}")
        if os.environ.get("SETTLEMENT_DEBUG") == "1":
            traceback.print_exc()
        else:
            hint("如需详细堆栈，请设置环境变量 SETTLEMENT_DEBUG=1 后重试。")
        return 1

    if result.output is not None:
        ok(f"输出文件：{result.output}")
    if result.preview is not None:
        info(f"数据预览：{result.preview.path.name}")
    if result.comparison_report is not None:
        info(f"差异报告：{result.comparison_report.name}")
    layout = "排版 自动分页" if result.pages is None else f"排版 {result.pages} 页"
    info(
        f"数据行 {result.rows} 行；{layout}"
        + (f"，补空行 {result.padded_rows - result.rows} 行" if result.padded_rows > result.rows else "")
    )
    if result.validation_report is not None:
        info(f"数据检查报告：{result.validation_report.name}")
    if result.run_manifest is not None:
        info(f"运行记录：{result.run_manifest}")
    if result.backup is not None:
        info(f"旧文件备份：{result.backup}")
    ok("预览完成" if result.preview is not None else "全部完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
