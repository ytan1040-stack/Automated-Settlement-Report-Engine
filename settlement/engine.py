# -*- coding: utf-8 -*-
"""计算引擎：把源数据行转换成可直接渲染的结算行。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from .config import BillingItem, RuleConfig
from .constants import (
    BASIS_FIXED,
    BASIS_PER_TIME,
    FIELD_BASE_AMOUNT,
    FIELD_DETAIL,
    FIELD_DETAIL_TOTAL,
    FIELD_FIXED_AMOUNT,
    FIELD_FIXED_BASE,
    FIELD_FLOATING_AMOUNT,
    FIELD_FLOATING_BASE,
    FIELD_TOTAL,
)
from .data import DataTable
from .errors import DataError
from .utils import (
    canonical_field,
    fmt_amount,
    fmt_number,
    is_blank,
    norm_header,
    quantize,
    to_decimal,
    to_text,
)


@dataclass
class ItemResult:
    item: BillingItem
    quantity: Decimal          # 实际计费数量（按次为 0/1，按数量为原数量，固定金额为 1）
    amount: Decimal
    text: str
    column_present: bool


@dataclass
class RowCalculation:
    record: dict[str, Any]
    numbers: dict[str, Decimal] = field(default_factory=dict)


@dataclass
class CalculationResult:
    rows: list[RowCalculation]
    warnings: list[str] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.rows)


_TRUE_MARKERS = {
    "是",
    "有",
    "√",
    "✓",
    "✔",
    "y",
    "yes",
    "true",
    "到",
    "出勤",
}
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


def calculate(config: RuleConfig, table: DataTable) -> CalculationResult:
    """遍历数据行，计算基本工资、每项计费金额、明细文字与应付合计。"""
    base = config.base
    decimals = base.decimals
    warnings: list[str] = []
    warned: set[str] = set()

    first_source = table.records[0] if table.records else {}
    item_fields = {item.name: _resolve_field(first_source, item.name) for item in config.items}
    for item in config.items:
        if item.basis == BASIS_FIXED:
            continue
        if item_fields[item.name] is None:
            message = (
                f"源数据缺少计费项列「{item.name}」，该项按 0 处理。"
                "如果这是笔误，请修改「规则配置.xlsx → 动态计费项」或「源数据.xlsx」的列名。"
            )
            if message not in warned:
                warnings.append(message)
                warned.add(message)

    fixed_field = _resolve_field(first_source, base.fixed_base_column)
    floating_field = _resolve_field(first_source, base.floating_base_column)
    if fixed_field is None and base.fixed_base_price == 0:
        warnings.append(
            f"源数据缺少固定基数列「{base.fixed_base_column}」，且固定基数单价为 0，按 0 处理。"
        )
    if floating_field is None and base.floating_base_price == 0:
        warnings.append(
            f"源数据缺少浮动基数列「{base.floating_base_column}」，且浮动基数单价为 0，按 0 处理。"
        )

    rows: list[RowCalculation] = []
    for index, source in enumerate(table.records):
        excel_row = table.header_row + index + 1
        record = dict(source)

        fixed_base = _read_number(
            source,
            fixed_field,
            base.fixed_base_column,
            table.path,
            excel_row,
            required=base.fixed_base_price != 0,
        )
        floating_base = _read_number(
            source,
            floating_field,
            base.floating_base_column,
            table.path,
            excel_row,
            required=base.floating_base_price != 0,
        )

        fixed_amount = quantize(base.fixed_base_price * fixed_base, decimals)
        floating_amount = quantize(base.floating_base_price * floating_base, decimals)
        basic_amount = quantize(fixed_amount + floating_amount, decimals)

        item_results: list[ItemResult] = []
        for item in config.items:
            field_name = item_fields[item.name]
            quantity = _read_quantity(
                source,
                field_name,
                item,
                table.path,
                excel_row,
            )
            if item.basis == BASIS_FIXED:
                amount = quantize(item.price, decimals)
                quantity = Decimal(1)
            elif item.basis == BASIS_PER_TIME:
                amount = quantize(item.price, decimals) if quantity != 0 else Decimal(0)
            else:
                amount = quantize(item.price * quantity, decimals)

            text = _format_item_text(item, quantity, amount, base.item_format, decimals)
            item_results.append(
                ItemResult(
                    item=item,
                    quantity=quantity,
                    amount=amount,
                    text=text,
                    column_present=field_name is not None,
                )
            )

        detail_total = quantize(sum((result.amount for result in item_results), Decimal(0)), decimals)
        visible_texts = [result.text for result in item_results if result.amount != 0 and result.text]
        detail_text = _join_detail(visible_texts, base.separator, base.total_prefix, detail_total, decimals)
        total_amount = quantize(basic_amount + detail_total, decimals)

        numbers: dict[str, Decimal] = {
            FIELD_FIXED_BASE: fixed_base,
            FIELD_FLOATING_BASE: floating_base,
            FIELD_FIXED_AMOUNT: fixed_amount,
            FIELD_FLOATING_AMOUNT: floating_amount,
            FIELD_BASE_AMOUNT: basic_amount,
            FIELD_DETAIL_TOTAL: detail_total,
            FIELD_TOTAL: total_amount,
        }

        # 保留源数据中其他数字列，方便模板表尾对自定义数字列求和。
        for field_name, value in source.items():
            if field_name in numbers or is_blank(value):
                continue
            try:
                numbers[field_name] = to_decimal(value)
            except ValueError:
                pass

        if fixed_field is None:
            record[FIELD_FIXED_BASE] = fmt_number(fixed_base)
        if floating_field is None:
            record[FIELD_FLOATING_BASE] = fmt_number(floating_base)

        record[FIELD_FIXED_AMOUNT] = fmt_amount(fixed_amount, decimals)
        record[FIELD_FLOATING_AMOUNT] = fmt_amount(floating_amount, decimals)
        record[FIELD_BASE_AMOUNT] = fmt_amount(basic_amount, decimals)
        record[FIELD_DETAIL] = detail_text
        record[FIELD_DETAIL_TOTAL] = fmt_amount(detail_total, decimals)
        record[FIELD_TOTAL] = fmt_amount(total_amount, decimals)

        for result in item_results:
            record[result.item.name] = fmt_amount(result.amount, decimals)
            quantity_field = f"{result.item.name}数量"
            if quantity_field not in record:
                record[quantity_field] = fmt_number(result.quantity)
            numbers[result.item.name] = result.amount
            numbers[quantity_field] = result.quantity

        rows.append(RowCalculation(record=record, numbers=numbers))

    return CalculationResult(rows=rows, warnings=_dedupe(warnings))


# ---------------------------------------------------------------------------
# 字段与数量解析
# ---------------------------------------------------------------------------
def _resolve_field(record: dict[str, Any], name: str) -> str | None:
    if not name:
        return None
    if name in record:
        return name
    normalized = norm_header(name)
    if normalized in record:
        return normalized
    alias = canonical_field(name)
    if alias in record:
        return alias
    return None


def _read_number(
    record: dict[str, Any],
    field_name: str | None,
    label: str,
    path: Path,
    excel_row: int,
    *,
    required: bool,
) -> Decimal:
    if field_name is None:
        if required:
            raise DataError(
                f"「{path.name}」缺少基数列「{label}」。"
                "请检查列名，或在「规则配置.xlsx → 字段映射」中把源数据列映射为「固定基数」/「浮动基数」。"
            )
        return Decimal(0)

    raw = record.get(field_name)
    if is_blank(raw):
        return Decimal(0)
    try:
        return to_decimal(raw)
    except ValueError as exc:
        raise DataError(
            f"「{path.name}」第 {excel_row} 行「{label}」不是有效数字：{to_text(raw)!r}。"
        ) from exc


def _read_quantity(
    record: dict[str, Any],
    field_name: str | None,
    item: BillingItem,
    path: Path,
    excel_row: int,
) -> Decimal:
    if item.basis == BASIS_FIXED:
        return Decimal(1)
    if field_name is None:
        return Decimal(0)

    raw = record.get(field_name)
    if is_blank(raw):
        return Decimal(0)

    text = to_text(raw)
    lowered = text.strip().lower()
    if item.basis == BASIS_PER_TIME and lowered in _TRUE_MARKERS:
        return Decimal(1)
    if lowered in _FALSE_MARKERS:
        return Decimal(0)

    cleaned = _QUANTITY_SUFFIX_RE.sub("", text).strip()
    if not cleaned:
        return Decimal(0)
    try:
        quantity = to_decimal(cleaned)
    except ValueError as exc:
        raise DataError(
            f"「{path.name}」第 {excel_row} 行计费项「{item.name}」不是有效数量：{text!r}。"
            f"「{item.basis}」请填写数字"
            + ("，或填写“是/否”。" if item.basis == BASIS_PER_TIME else "。")
        ) from exc

    if item.basis == BASIS_PER_TIME:
        return Decimal(1) if quantity != 0 else Decimal(0)
    return quantity


# ---------------------------------------------------------------------------
# 明细文字
# ---------------------------------------------------------------------------
_SIGN_PLUS_BEFORE_NEGATIVE_RE = re.compile(r"(?:\+\s*)+(?=-\d)")


def _format_item_text(
    item: BillingItem,
    quantity: Decimal,
    amount: Decimal,
    item_format: str,
    decimals: int,
) -> str:
    try:
        return item_format.format(
            name=item.name,
            amount=fmt_amount(amount, decimals),
            qty=fmt_number(quantity),
            price=fmt_amount(item.price, decimals),
        ).strip()
    except Exception as exc:  # noqa: BLE001 - 配置已校验，这里只做最后兜底
        raise DataError(f"计费项「{item.name}」的展示格式无法套用：{exc}") from exc


def _join_detail(
    texts: Iterable[str],
    separator: str,
    total_prefix: str,
    detail_total: Decimal,
    decimals: int,
) -> str:
    visible = [text for text in texts if text]
    if not visible:
        return ""
    body = separator.join(visible)
    if total_prefix:
        body += total_prefix + fmt_amount(detail_total, decimals)
    # 配置里的「+」是给正数用的；负数会形成「+-24.00」。
    # 在最终文本上统一去掉负数前的加号，兼容 {name}+{amount}、
    # 「, 共+」等默认/自定义格式，同时不影响正数。
    return _SIGN_PLUS_BEFORE_NEGATIVE_RE.sub("", body)


def _dedupe(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result
