# -*- coding: utf-8 -*-
"""数据预览：不生成 Word，先把「数据表 + 合计 + 页数」写出来。

预览主要解决两个问题：

1. 用户不想为了看一行数据就打开 Word；
2. 生成前先确认行数、页数、合计是否和预期一致。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import RuleConfig
from .constants import (
    FIELD_BASE_AMOUNT,
    FIELD_DEPARTMENT,
    FIELD_DETAIL,
    FIELD_DETAIL_TOTAL,
    FIELD_EMPLOYEE_ID,
    FIELD_FIXED_AMOUNT,
    FIELD_FIXED_BASE,
    FIELD_FLOATING_AMOUNT,
    FIELD_FLOATING_BASE,
    FIELD_NAME,
    FIELD_REMARK,
    FIELD_SEQ,
    FIELD_TOTAL,
    NON_TOTAL_FIELDS,
    PREVIEW_FILE,
)
from .data import DataTable
from .engine import CalculationResult
from .excel import save_workbook_atomic
from .utils import emit, fmt_number, info, to_text


@dataclass
class PreviewResult:
    path: Path
    rows: int
    columns: list[str]
    pages: int | None
    totals: dict[str, Decimal] = field(default_factory=dict)
    samples: list[list[str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def write_preview_report(
    root: Path,
    data: DataTable,
    calculation: CalculationResult,
    config: RuleConfig,
    *,
    pages: int | None,
    padding: int = 0,
    rows_per_page: int | None = None,
) -> PreviewResult:
    """写出 ``数据预览.xlsx``，返回预览结果。"""
    root = Path(root)
    columns = _preview_columns(config, calculation)
    totals = _collect_totals(calculation)
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "数据预览"
    _write_data_sheet(
        sheet,
        data,
        calculation,
        columns,
        totals,
        config,
        pages=pages,
        padding=padding,
        rows_per_page=rows_per_page,
    )

    summary = workbook.create_sheet("汇总")
    _write_summary_sheet(
        summary,
        data,
        calculation,
        config,
        columns,
        totals,
        pages=pages,
        padding=padding,
        rows_per_page=rows_per_page,
    )

    path = root / PREVIEW_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    save_workbook_atomic(workbook, path)
    samples: list[list[str]] = []
    for row in calculation.rows[:20]:
        samples.append([to_text(row.record.get(name, "")) for name in columns])
    return PreviewResult(
        path=path,
        rows=len(calculation.rows),
        columns=columns,
        pages=pages,
        totals=totals,
        samples=samples,
    )


def print_preview(result: PreviewResult, *, limit: int = 10) -> None:
    """在控制台打印预览摘要和前几行，避免用户必须打开 Excel。"""
    page_text = "自动分页" if result.pages is None else f"{result.pages} 页"
    info(
        f"数据预览：{result.rows} 行 × {len(result.columns)} 列，"
        f"{page_text}，已写入「{result.path.name}」"
    )

    summary = _totals_summary(result.totals)
    if summary:
        info(f"预览合计：{summary}")

    if not result.rows:
        return

    # 控制台只展示最有用的几列；完整数据在 Excel 中。
    preferred = [FIELD_SEQ, FIELD_NAME, FIELD_EMPLOYEE_ID, FIELD_DETAIL, FIELD_TOTAL]
    indexes = [
        (result.columns.index(name), name)
        for name in preferred
        if name in result.columns
    ]
    if not indexes:
        indexes = list(enumerate(result.columns[:8]))

    emit("  " + " | ".join(_short(name, 12) for _, name in indexes))
    emit("  " + "-+-".join("-" * 12 for _ in indexes))
    for values in result.samples[:limit]:
        emit("  " + " | ".join(_short(values[index], 12) for index, _ in indexes))

    if result.rows > limit:
        info(f"仅显示前 {limit} 行；完整 {result.rows} 行请打开「{result.path.name}」查看。")


def _preview_columns(config: RuleConfig, calculation: CalculationResult) -> list[str]:
    """按业务阅读顺序排列预览列，并补上源数据里的自定义列。"""
    base = config.base
    priority = [
        FIELD_SEQ,
        FIELD_NAME,
        FIELD_DEPARTMENT,
        FIELD_EMPLOYEE_ID,
        base.fixed_base_column,
        base.floating_base_column,
    ]
    for item in config.items:
        priority.append(item.name)
        priority.append(f"{item.name}数量")
    priority.extend(
        [
            FIELD_FIXED_AMOUNT,
            FIELD_FLOATING_AMOUNT,
            FIELD_BASE_AMOUNT,
            FIELD_DETAIL,
            FIELD_DETAIL_TOTAL,
            FIELD_TOTAL,
            FIELD_REMARK,
        ]
    )

    present: list[str] = []
    known: set[str] = set()
    for row in calculation.rows:
        for key in row.record:
            known.add(key)

    for name in priority:
        if name in known and name not in present:
            present.append(name)

    # 源数据中的自定义列（例如“项目”“月份”）放在备注之前，保留可读性。
    extras = sorted(known - set(present))
    if FIELD_REMARK in present:
        index = present.index(FIELD_REMARK)
        present[index:index] = extras
    else:
        present.extend(extras)
    return present


def _collect_totals(calculation: CalculationResult) -> dict[str, Decimal]:
    totals: dict[str, Decimal] = {}
    for row in calculation.rows:
        for name, value in row.numbers.items():
            if name in NON_TOTAL_FIELDS:
                continue
            totals[name] = totals.get(name, Decimal(0)) + value
    return totals


def _write_data_sheet(
    sheet,
    data: DataTable,
    calculation: CalculationResult,
    columns: list[str],
    totals: dict[str, Decimal],
    config: RuleConfig,
    *,
    pages: int | None,
    padding: int,
    rows_per_page: int | None,
) -> None:
    title_font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    header_font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(columns) + 1))
    sheet.cell(1, 1, "数据预览").font = title_font
    sheet.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
    sheet.row_dimensions[1].height = 26

    page_text = "自动分页" if pages is None else f"{pages} 页"
    sheet.cell(
        2,
        1,
        f"源文件：{data.path.name}  |  工作表：{data.sheet_name}  |  "
        f"数据 {len(calculation.rows)} 行  |  {page_text}"
        + (f"  |  补空行 {padding} 行" if padding else ""),
    ).font = Font(name="微软雅黑", size=9, color="666666")

    header_row = 4
    headers = ["页码"] + columns
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(header_row, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    sheet.row_dimensions[header_row].height = 28
    sheet.freeze_panes = f"A{header_row + 1}"

    for index, row in enumerate(calculation.rows, start=header_row + 1):
        if rows_per_page and rows_per_page > 0:
            page_number = index - header_row  # 1-based data index
            page_number = (page_number - 1) // rows_per_page + 1
        else:
            page_number = "自动"
        sheet.cell(index, 1, page_number).alignment = Alignment(horizontal="center")
        sheet.cell(index, 1).border = border
        for column, name in enumerate(columns, start=2):
            value = to_text(row.record.get(name, ""))
            cell = sheet.cell(index, column, value)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(name == FIELD_DETAIL))

    # 表尾合计：只对数字列写值。
    total_row = header_row + 1 + len(calculation.rows)
    sheet.cell(total_row, 1, "合计").font = Font(name="微软雅黑", bold=True)
    sheet.cell(total_row, 1).fill = total_fill
    sheet.cell(total_row, 1).border = border
    for column, name in enumerate(columns, start=2):
        value = totals.get(name)
        cell = sheet.cell(
            total_row,
            column,
            fmt_number(value, 6) if value is not None else "",
        )
        cell.font = Font(name="微软雅黑", bold=True)
        cell.fill = total_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="right" if value is not None else "center")

    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{total_row - 1}"

    sheet.column_dimensions["A"].width = 7
    for column, name in enumerate(columns, start=2):
        width = 8
        if name in (FIELD_NAME, FIELD_DEPARTMENT):
            width = 12
        elif name == FIELD_DETAIL:
            width = 48
        elif name in (FIELD_TOTAL, FIELD_DETAIL_TOTAL):
            width = 13
        elif "数量" in name:
            width = 9
        sheet.column_dimensions[get_column_letter(column)].width = width


def _write_summary_sheet(
    sheet,
    data: DataTable,
    calculation: CalculationResult,
    config: RuleConfig,
    columns: list[str],
    totals: dict[str, Decimal],
    *,
    pages: int | None,
    padding: int,
    rows_per_page: int | None,
) -> None:
    sheet.merge_cells("A1:C1")
    sheet["A1"] = "预览汇总"
    sheet["A1"].font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    sheet.row_dimensions[1].height = 26

    rows: list[tuple[str, Any]] = [
        ("源文件", data.path.name),
        ("工作表", data.sheet_name),
        ("数据行数", len(calculation.rows)),
        ("输出列数", len(columns)),
        ("每页固定行数", rows_per_page if rows_per_page else "自动分页"),
        ("页数", pages if pages is not None else "自动"),
        ("补空行", padding),
        ("固定基数单价", fmt_number(config.base.fixed_base_price, 6)),
        ("浮动基数单价", fmt_number(config.base.floating_base_price, 6)),
    ]
    for offset, (label, value) in enumerate(rows, start=3):
        sheet.cell(offset, 1, label).font = Font(name="微软雅黑", bold=True)
        sheet.cell(offset, 2, value)

    row = 3 + len(rows) + 1
    sheet.cell(row, 1, "数值字段合计").font = Font(name="微软雅黑", bold=True)
    row += 1
    if totals:
        for name in _total_order(config, totals):
            sheet.cell(row, 1, name)
            sheet.cell(row, 2, float(totals[name]))
            row += 1
    else:
        sheet.cell(row, 1, "（没有可求和的数值列）")

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 42
    sheet.freeze_panes = "A3"


def _total_order(config: RuleConfig, totals: dict[str, Decimal]) -> list[str]:
    preferred = [
        FIELD_FIXED_BASE,
        FIELD_FLOATING_BASE,
        FIELD_FIXED_AMOUNT,
        FIELD_FLOATING_AMOUNT,
        FIELD_BASE_AMOUNT,
        FIELD_DETAIL_TOTAL,
        FIELD_TOTAL,
    ]
    for item in config.items:
        preferred.append(item.name)
        preferred.append(f"{item.name}数量")
    ordered = [name for name in preferred if name in totals]
    ordered.extend(sorted(name for name in totals if name not in ordered))
    return ordered


def _totals_summary(totals: dict[str, Decimal]) -> str:
    keys = [FIELD_TOTAL, FIELD_BASE_AMOUNT, FIELD_DETAIL_TOTAL]
    parts = [f"{name} {fmt_number(totals[name], 6)}" for name in keys if name in totals]
    if not parts:
        return ""
    return "；".join(parts)


def _short(text: str, width: int) -> str:
    text = to_text(text)
    return text if len(text) <= width else text[: width - 1] + "…"
