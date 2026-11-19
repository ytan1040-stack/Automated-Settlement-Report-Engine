# -*- coding: utf-8 -*-
"""数据校验与「数据检查报告.xlsx」生成。

校验目标是“在算错之前发现问题”，而不是替代 Excel 的数据验证。
它只做保守检查：不把可疑但合法的值直接判死，能警告就警告。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import BillingItem, RuleConfig
from .constants import (
    BASIS_FIXED,
    BASIS_PER_QTY,
    BASIS_PER_TIME,
    FIELD_EMPLOYEE_ID,
    FIELD_NAME,
    FIELD_SEQ,
    NON_TOTAL_FIELDS,
)
from .data import DataTable
from .excel import save_workbook_atomic
from .utils import canonical_field, fmt_number, is_blank, norm_header, to_decimal, to_text

_TRUE_MARKERS = {"是", "有", "√", "✓", "✔", "y", "yes", "true", "到", "出勤"}
_FALSE_MARKERS = {
    "",
    "否",
    "无",
    "×",
    "x",
    "n",
    "no",
    "false",
    "0",
    "-",
    "—",
    "/",
    "none",
    "null",
    "n/a",
}
_QUANTITY_SUFFIX_RE = re.compile(r"\s*(次|个|份|人|天|小时|项|笔|台|套|张|间)$")


@dataclass(frozen=True)
class ValidationIssue:
    level: str          # "错误" / "警告"
    code: str
    message: str
    row: int | None = None
    column: str = ""


@dataclass
class ValidationResult:
    path: Path
    sheet_name: str
    rows: int
    issues: list[ValidationIssue] = field(default_factory=list)
    control_totals: dict[str, Decimal] = field(default_factory=dict)
    checked_at: datetime = field(default_factory=datetime.now)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.level == "错误"]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.level == "警告"]

    @property
    def ok(self) -> bool:
        return not self.errors

    @property
    def status(self) -> str:
        if self.errors:
            return "未通过"
        if self.warnings:
            return "通过（有警告）"
        return "通过"


def validate_data(table: DataTable, config: RuleConfig) -> ValidationResult:
    """检查基数列、计费项、重复编号、数量与金额一致性。"""
    result = ValidationResult(
        path=table.path,
        sheet_name=table.sheet_name,
        rows=table.row_count,
    )
    fields = set(table.fields)
    records = table.records

    _check_base_columns(table, config, fields, result)
    _check_items(table, config, fields, result)
    _check_identifiers(table, fields, result)
    _check_qty_price_amount(table, fields, result)
    _collect_control_totals(table, result)
    return result


# ---------------------------------------------------------------------------
# 单项检查
# ---------------------------------------------------------------------------
def _check_base_columns(
    table: DataTable,
    config: RuleConfig,
    fields: set[str],
    result: ValidationResult,
) -> None:
    base = config.base
    for label, column, price in (
        ("固定基数", base.fixed_base_column, base.fixed_base_price),
        ("浮动基数", base.floating_base_column, base.floating_base_price),
    ):
        field_name = _resolve_field(fields, column)
        if field_name is None:
            if price != 0:
                result.issues.append(
                    ValidationIssue(
                        "错误",
                        "BASE_COLUMN_MISSING",
                        f"源数据缺少基数列「{column}」，但{label}单价为 {fmt_number(price)}。",
                        column=column,
                    )
                )
            continue
        if price == 0:
            continue
        for index, record in enumerate(table.records):
            raw = record.get(field_name)
            excel_row = table.header_row + index + 1
            if is_blank(raw):
                result.issues.append(
                    ValidationIssue(
                        "错误",
                        "BASE_VALUE_EMPTY",
                        f"{label}为空；{label}单价不为 0 时必须有值。",
                        row=excel_row,
                        column=column,
                    )
                )
                continue
            try:
                to_decimal(raw)
            except ValueError:
                result.issues.append(
                    ValidationIssue(
                        "错误",
                        "BASE_VALUE_INVALID",
                        f"{label}不是有效数字：{to_text(raw)!r}。",
                        row=excel_row,
                        column=column,
                    )
                )


def _check_items(
    table: DataTable,
    config: RuleConfig,
    fields: set[str],
    result: ValidationResult,
) -> None:
    for item in config.items:
        if item.basis == BASIS_FIXED:
            continue
        field_name = _resolve_field(fields, item.name)
        if field_name is None:
            if item.price != 0:
                result.issues.append(
                    ValidationIssue(
                        "警告",
                        "ITEM_COLUMN_MISSING",
                        f"源数据缺少计费项列「{item.name}」，该项会按 0 计算。"
                        "如果这是笔误，请修改动态计费项或源数据列名。",
                        column=item.name,
                    )
                )
            continue

        for index, record in enumerate(table.records):
            raw = record.get(field_name)
            if is_blank(raw):
                continue
            excel_row = table.header_row + index + 1
            try:
                quantity = _parse_quantity(raw, item)
            except ValueError:
                expect = "数字或 是/否" if item.basis == BASIS_PER_TIME else "数字"
                result.issues.append(
                    ValidationIssue(
                        "错误",
                        "ITEM_VALUE_INVALID",
                        f"计费项「{item.name}」的值不是有效{expect}：{to_text(raw)!r}。",
                        row=excel_row,
                        column=item.name,
                    )
                )
                continue
            if item.basis == BASIS_PER_QTY and quantity < 0:
                result.issues.append(
                    ValidationIssue(
                        "警告",
                        "ITEM_NEGATIVE",
                        f"计费项「{item.name}」数量为负数（{fmt_number(quantity)}）；"
                        "请确认这是退款/冲销，而不是录入错误。",
                        row=excel_row,
                        column=item.name,
                    )
                )


def _check_identifiers(
    table: DataTable,
    fields: set[str],
    result: ValidationResult,
) -> None:
    for field_name in (FIELD_EMPLOYEE_ID, FIELD_SEQ, "编号"):
        if field_name not in fields:
            continue
        seen: dict[str, int] = {}
        for index, record in enumerate(table.records):
            value = to_text(record.get(field_name)).strip()
            if not value:
                continue
            excel_row = table.header_row + index + 1
            if value in seen:
                result.issues.append(
                    ValidationIssue(
                        "警告",
                        "DUPLICATE_ID",
                        f"「{field_name}」重复：{value!r}，首次出现在第 {seen[value]} 行。",
                        row=excel_row,
                        column=field_name,
                    )
                )
            else:
                seen[value] = excel_row


def _check_qty_price_amount(
    table: DataTable,
    fields: set[str],
    result: ValidationResult,
) -> None:
    """如果同时存在 数量/单价/金额，检查三者是否大致一致。"""
    qty_field = _resolve_field(fields, "数量")
    price_field = _resolve_field(fields, "单价")
    amount_field = _resolve_field(fields, "金额")
    if not (qty_field and price_field and amount_field):
        return

    for index, record in enumerate(table.records):
        qty_raw = record.get(qty_field)
        price_raw = record.get(price_field)
        amount_raw = record.get(amount_field)
        if any(is_blank(value) for value in (qty_raw, price_raw, amount_raw)):
            continue
        try:
            qty = to_decimal(qty_raw)
            price = to_decimal(price_raw)
            amount = to_decimal(amount_raw)
        except ValueError:
            continue
        expected = qty * price
        if abs(expected - amount) > Decimal("0.011"):
            result.issues.append(
                ValidationIssue(
                    "警告",
                    "QTY_PRICE_AMOUNT_MISMATCH",
                    f"数量 × 单价 = {fmt_number(expected)}，与金额 {fmt_number(amount)} 不一致。",
                    row=table.header_row + index + 1,
                    column="金额",
                )
            )


def _collect_control_totals(table: DataTable, result: ValidationResult) -> None:
    """为报告提供一个“源数据数值列合计”，方便和财务总数对账。"""
    numeric_fields: set[str] = set()
    for record in table.records:
        for field_name, value in record.items():
            if field_name in NON_TOTAL_FIELDS or field_name in numeric_fields:
                continue
            if is_blank(value):
                continue
            try:
                to_decimal(value)
            except ValueError:
                continue
            numeric_fields.add(field_name)

    for field_name in sorted(numeric_fields):
        total = Decimal(0)
        for record in table.records:
            try:
                total += to_decimal(record.get(field_name))
            except (TypeError, ValueError):
                pass
        result.control_totals[field_name] = total


# ---------------------------------------------------------------------------
# 报告
# ---------------------------------------------------------------------------
def write_validation_report(result: ValidationResult, path: Path) -> Path:
    """写出数据检查报告；返回报告路径。"""
    path = Path(path)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "检查结果"

    title_font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet.merge_cells("A1:D1")
    sheet["A1"] = "数据检查报告"
    sheet["A1"].font = title_font
    sheet["A1"].alignment = Alignment(horizontal="left", vertical="center")
    sheet.row_dimensions[1].height = 26

    rows: list[tuple[str, Any]] = [
        ("源文件", result.path.name),
        ("工作表", result.sheet_name),
        ("数据行数", result.rows),
        ("检查时间", result.checked_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("结论", result.status),
        ("错误数", len(result.errors)),
        ("警告数", len(result.warnings)),
    ]
    for offset, (label, value) in enumerate(rows, start=3):
        sheet.cell(offset, 1, label).font = Font(name="微软雅黑", bold=True)
        sheet.cell(offset, 2, value)

    row = 11
    sheet.cell(row, 1, "源数据数值列合计").font = Font(name="微软雅黑", bold=True)
    row += 1
    if result.control_totals:
        for name, total in result.control_totals.items():
            sheet.cell(row, 1, name)
            sheet.cell(row, 2, float(total))
            row += 1
    else:
        sheet.cell(row, 1, "（没有可求和的数值列）")
        row += 1

    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 42
    sheet.freeze_panes = "A3"

    issues = workbook.create_sheet("问题清单")
    headers = ["级别", "代码", "行号", "列", "说明"]
    widths = [8, 28, 8, 18, 80]
    for column, (header, width) in enumerate(zip(headers, widths), start=1):
        cell = issues.cell(1, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        issues.column_dimensions[get_column_letter(column)].width = width
    issues.row_dimensions[1].height = 22
    issues.freeze_panes = "A2"

    if not result.issues:
        issues.cell(2, 1, "无")
        issues.cell(2, 5, "没有发现需要处理的问题。")
    else:
        for offset, issue in enumerate(result.issues, start=2):
            values = [
                issue.level,
                issue.code,
                issue.row if issue.row is not None else "",
                issue.column,
                issue.message,
            ]
            for column, value in enumerate(values, start=1):
                cell = issues.cell(offset, column, value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=(column == 5))
            if issue.level == "错误":
                issues.cell(offset, 1).font = Font(color="C00000", bold=True)
            else:
                issues.cell(offset, 1).font = Font(color="BF8F00", bold=True)

    path.parent.mkdir(parents=True, exist_ok=True)
    save_workbook_atomic(workbook, path)
    return path


# ---------------------------------------------------------------------------
# 辅助
# ---------------------------------------------------------------------------
def _resolve_field(fields: set[str], name: str) -> str | None:
    if not name:
        return None
    if name in fields:
        return name
    normalized = norm_header(name)
    if normalized in fields:
        return normalized
    alias = canonical_field(name)
    if alias in fields:
        return alias
    return None


def _parse_quantity(raw: Any, item: BillingItem) -> Decimal:
    text = to_text(raw)
    lowered = text.strip().lower()
    if item.basis == BASIS_PER_TIME and lowered in _TRUE_MARKERS:
        return Decimal(1)
    if lowered in _FALSE_MARKERS:
        return Decimal(0)
    cleaned = _QUANTITY_SUFFIX_RE.sub("", text).strip()
    if not cleaned:
        return Decimal(0)
    quantity = to_decimal(cleaned)
    if item.basis == BASIS_PER_TIME:
        return Decimal(1) if quantity != 0 else Decimal(0)
    return quantity
