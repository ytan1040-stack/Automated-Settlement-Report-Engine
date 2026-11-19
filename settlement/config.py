# -*- coding: utf-8 -*-
"""读取并校验 规则配置.xlsx。"""

from __future__ import annotations

import re
import string
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .constants import (
    BASIS_OPTIONS,
    CONFIG_FILE,
    DEFAULT_DECIMALS,
    DEFAULT_FIXED_COLUMN,
    DEFAULT_FLOATING_COLUMN,
    DEFAULT_FONT_AUTO,
    DEFAULT_HEADER_ROWS,
    DEFAULT_ITEM_FORMAT,
    DEFAULT_OUTPUT_PATTERN,
    DEFAULT_REPORT_TITLE,
    DEFAULT_ROWS_PER_PAGE,
    DEFAULT_SEPARATOR,
    DEFAULT_TOTAL_PREFIX,
    KEY_DECIMALS,
    KEY_FIXED_COLUMN,
    KEY_FIXED_PRICE,
    KEY_FLOATING_COLUMN,
    KEY_FLOATING_PRICE,
    KEY_FONT_ASCII,
    KEY_FONT_EAST_ASIA,
    KEY_FONT_SIZE,
    KEY_HEADER_ROWS,
    KEY_ITEM_FORMAT,
    KEY_OUTPUT_NAME,
    KEY_REPORT_TITLE,
    KEY_ROWS_PER_PAGE,
    KEY_SEPARATOR,
    KEY_TABLE_INDEX,
    KEY_TOTAL_PREFIX,
    REQUIRED_BASE_KEYS,
    RESERVED_FIELDS,
    SHEET_BASE,
    SHEET_ITEMS,
    SHEET_MAPPING,
)
from .errors import ConfigError
from .excel import find_header_row, find_sheet, load_workbook_safe, require_columns
from .utils import canonical_field, is_blank, norm_header, to_decimal, to_text


@dataclass(frozen=True)
class BaseConfig:
    fixed_base_price: Decimal
    floating_base_price: Decimal
    item_format: str
    separator: str
    total_prefix: str
    rows_per_page: int | None
    fixed_base_column: str
    floating_base_column: str
    report_title: str
    decimals: int
    output_pattern: str
    table_index: int | None
    header_rows: int
    font_east_asia: str = DEFAULT_FONT_AUTO
    font_ascii: str = DEFAULT_FONT_AUTO
    font_size_pt: float | None = None


@dataclass(frozen=True)
class BillingItem:
    name: str
    price: Decimal
    basis: str
    note: str = ""


@dataclass
class RuleConfig:
    path: Path
    base: BaseConfig
    items: tuple[BillingItem, ...]
    mapping: dict[str, str]
    warnings: tuple[str, ...] = field(default_factory=tuple)


# ---------------------------------------------------------------------------
# 公开入口
# ---------------------------------------------------------------------------
def load_config(path: Path) -> RuleConfig:
    path = Path(path)
    if not path.exists():
        raise ConfigError(
            f"未找到规则配置「{CONFIG_FILE}」。"
            "请先运行 init_config.py 生成配置模板，或直接双击 一键运行.bat。"
        )

    workbook = load_workbook_safe(path)
    try:
        ws_base = find_sheet(workbook, SHEET_BASE, 0)
        ws_items = find_sheet(workbook, SHEET_ITEMS, 1)
        ws_mapping = find_sheet(workbook, SHEET_MAPPING, 2)

        warnings: list[str] = []
        base = _load_base(ws_base, warnings)
        items = _load_items(ws_items, warnings)
        mapping = _load_mapping(ws_mapping, warnings)
    finally:
        workbook.close()

    return RuleConfig(
        path=path,
        base=base,
        items=tuple(items),
        mapping=mapping,
        warnings=tuple(warnings),
    )


# ---------------------------------------------------------------------------
# Sheet1 基础配置
# ---------------------------------------------------------------------------
def _load_base(ws: Worksheet, warnings: list[str]) -> BaseConfig:
    header_row, columns = find_header_row(ws, ("配置项",), max_scan=10)
    if header_row is None:
        key_column, value_column, start_row = 1, 2, 1
    else:
        key_column = columns.get(norm_header("配置项"), 1)
        value_column = columns.get(norm_header("值"), 2)
        start_row = header_row + 1

    values: dict[str, Any] = {}
    for row in range(start_row, ws.max_row + 1):
        key = norm_header(ws.cell(row=row, column=key_column).value)
        if not key:
            continue
        if key in values:
            raise ConfigError(f"「{SHEET_BASE}」配置项「{key}」重复出现，请只保留一行。")
        values[key] = ws.cell(row=row, column=value_column).value

    missing = [key for key in REQUIRED_BASE_KEYS if key not in values]
    if missing:
        raise ConfigError(
            f"「{SHEET_BASE}」缺少必需配置项：{'、'.join(missing)}。"
            "请重新运行 init_config.py 生成完整配置。"
        )

    fixed_price = _parse_decimal(values[KEY_FIXED_PRICE], KEY_FIXED_PRICE)
    floating_price = _parse_decimal(values[KEY_FLOATING_PRICE], KEY_FLOATING_PRICE)
    item_format = _parse_item_format(_text_or_default(values[KEY_ITEM_FORMAT], DEFAULT_ITEM_FORMAT))
    separator = _text_or_default(values[KEY_SEPARATOR], DEFAULT_SEPARATOR, allow_disable=True)
    total_prefix = _text_or_default(values[KEY_TOTAL_PREFIX], DEFAULT_TOTAL_PREFIX, allow_disable=True)
    rows_per_page = _parse_rows_per_page(values[KEY_ROWS_PER_PAGE])
    fixed_column = _column_name(values.get(KEY_FIXED_COLUMN), DEFAULT_FIXED_COLUMN)
    floating_column = _column_name(values.get(KEY_FLOATING_COLUMN), DEFAULT_FLOATING_COLUMN)
    report_title = _text_or_default(values.get(KEY_REPORT_TITLE), DEFAULT_REPORT_TITLE)
    decimals = _parse_decimals(values.get(KEY_DECIMALS))
    output_pattern = _parse_output_pattern(values.get(KEY_OUTPUT_NAME), warnings)
    table_index = _parse_table_index(values.get(KEY_TABLE_INDEX))
    header_rows = _parse_header_rows(values.get(KEY_HEADER_ROWS))
    font_east_asia = _parse_font_name(values.get(KEY_FONT_EAST_ASIA))
    font_ascii = _parse_font_name(values.get(KEY_FONT_ASCII))
    font_size_pt = _parse_font_size(values.get(KEY_FONT_SIZE))

    if fixed_price < 0 or floating_price < 0:
        warnings.append("固定基数单价或浮动基数单价为负数，请确认这是扣款而不是录入错误。")
    if rows_per_page is not None and rows_per_page > 100:
        warnings.append(f"每页固定行数为 {rows_per_page}，数值较大，请确认排版是否合适。")

    return BaseConfig(
        fixed_base_price=fixed_price,
        floating_base_price=floating_price,
        item_format=item_format,
        separator=separator,
        total_prefix=total_prefix,
        rows_per_page=rows_per_page,
        fixed_base_column=fixed_column,
        floating_base_column=floating_column,
        report_title=report_title,
        decimals=decimals,
        output_pattern=output_pattern,
        table_index=table_index,
        header_rows=header_rows,
        font_east_asia=font_east_asia,
        font_ascii=font_ascii,
        font_size_pt=font_size_pt,
    )


# ---------------------------------------------------------------------------
# Sheet2 动态计费项
# ---------------------------------------------------------------------------
_ITEM_NAME_HEADERS = ("计费项名称", "项目名称", "计费项", "名称")
_ITEM_PRICE_HEADERS = ("单价", "计费单价", "价格")
_ITEM_BASIS_HEADERS = ("计费基数", "计费方式", "基数类型")
_ITEM_NOTE_HEADERS = ("说明", "备注")


def _load_items(ws: Worksheet, warnings: list[str]) -> list[BillingItem]:
    header_row, columns = find_header_row(
        ws,
        _ITEM_NAME_HEADERS + _ITEM_PRICE_HEADERS + _ITEM_BASIS_HEADERS,
        max_scan=10,
    )
    if header_row is None:
        raise ConfigError(
            f"「{SHEET_ITEMS}」找不到表头。"
            "请确保第 1 行包含：计费项名称、单价、计费基数、说明。"
        )

    name_column = _find_column(columns, _ITEM_NAME_HEADERS)
    price_column = _find_column(columns, _ITEM_PRICE_HEADERS)
    basis_column = _find_column(columns, _ITEM_BASIS_HEADERS)
    note_column = _find_column(columns, _ITEM_NOTE_HEADERS, required=False)
    missing = []
    if name_column is None:
        missing.append("计费项名称")
    if price_column is None:
        missing.append("单价")
    if basis_column is None:
        missing.append("计费基数")
    if missing:
        raise ConfigError(
            f"「{SHEET_ITEMS}」缺少必需的表头：{'、'.join(missing)}。"
            "请重新运行 init_config.py，或按模板补齐表头。"
        )

    items: list[BillingItem] = []
    seen_names: dict[str, int] = {}
    for row in range(header_row + 1, ws.max_row + 1):
        raw_name = ws.cell(row=row, column=name_column).value
        raw_price = ws.cell(row=row, column=price_column).value
        raw_basis = ws.cell(row=row, column=basis_column).value
        raw_note = ws.cell(row=row, column=note_column).value if note_column else None

        if is_blank(raw_name) and is_blank(raw_price) and is_blank(raw_basis) and is_blank(raw_note):
            continue
        if is_blank(raw_name):
            raise ConfigError(
                f"「{SHEET_ITEMS}」第 {row} 行填写了单价或计费基数，但计费项名称为空。"
            )

        name = norm_header(raw_name)
        normalized = name
        if normalized in seen_names:
            raise ConfigError(
                f"「{SHEET_ITEMS}」计费项「{name}」重复出现"
                f"（第 {seen_names[normalized]} 行与第 {row} 行），请合并为一行。"
            )
        seen_names[normalized] = row

        if name in RESERVED_FIELDS or canonical_field(name) in RESERVED_FIELDS:
            raise ConfigError(
                f"「{SHEET_ITEMS}」第 {row} 行：计费项名称「{name}」与系统字段重名。"
                "请换一个名称，例如「值班补贴」而不是「基本工资」。"
            )

        if is_blank(raw_price):
            raise ConfigError(f"「{SHEET_ITEMS}」第 {row} 行：计费项「{name}」的单价为空。")
        price = _parse_decimal(raw_price, f"第 {row} 行「{name}」的单价")

        basis = _WHITESPACE_RE.sub("", to_text(raw_basis))
        if not basis:
            raise ConfigError(
                f"「{SHEET_ITEMS}」第 {row} 行：计费项「{name}」的计费基数未填写。"
                f"可选值：{'、'.join(BASIS_OPTIONS)}。"
            )
        if basis not in BASIS_OPTIONS:
            raise ConfigError(
                f"「{SHEET_ITEMS}」第 {row} 行：计费基数「{basis}」无效。"
                f"请填写：{'、'.join(BASIS_OPTIONS)}。"
            )

        if price == 0:
            warnings.append(f"计费项「{name}」单价为 0，不会产生金额。")

        items.append(
            BillingItem(
                name=name,
                price=price,
                basis=basis,
                note=to_text(raw_note),
            )
        )

    if not items:
        warnings.append("「动态计费项」没有任何计费行，明细列将为空。")
    return items


# ---------------------------------------------------------------------------
# Sheet3 字段映射
# ---------------------------------------------------------------------------
_MAPPING_SOURCE_HEADERS = ("用户原始列名", "原始列名", "原列名", "源列名")
_MAPPING_TARGET_HEADERS = ("系统标准列名", "标准列名", "标准字段", "系统字段")


def _load_mapping(ws: Worksheet, warnings: list[str]) -> dict[str, str]:
    header_row, columns = find_header_row(
        ws,
        _MAPPING_SOURCE_HEADERS + _MAPPING_TARGET_HEADERS,
        max_scan=10,
    )
    if header_row is None:
        raise ConfigError(
            f"「{SHEET_MAPPING}」找不到表头。"
            "请确保第 1 行包含：用户原始列名、系统标准列名。"
        )

    source_column = _find_column(columns, _MAPPING_SOURCE_HEADERS)
    target_column = _find_column(columns, _MAPPING_TARGET_HEADERS)
    if source_column is None or target_column is None:
        raise ConfigError(
            f"「{SHEET_MAPPING}」表头必须同时包含"
            "「用户原始列名」和「系统标准列名」。"
        )

    mapping: dict[str, str] = {}
    for row in range(header_row + 1, ws.max_row + 1):
        raw_source = ws.cell(row=row, column=source_column).value
        raw_target = ws.cell(row=row, column=target_column).value
        if is_blank(raw_source) and is_blank(raw_target):
            continue
        if is_blank(raw_source):
            raise ConfigError(
                f"「{SHEET_MAPPING}」第 {row} 行：填写了系统标准列名，但用户原始列名为空。"
            )
        if is_blank(raw_target):
            raise ConfigError(
                f"「{SHEET_MAPPING}」第 {row} 行：用户原始列名"
                f"「{to_text(raw_source)}」没有对应的系统标准列名。"
            )
        source_key = norm_header(raw_source)
        target = norm_header(raw_target)
        if source_key in mapping:
            raise ConfigError(
                f"「{SHEET_MAPPING}」用户原始列名「{to_text(raw_source)}」重复出现，请只保留一行。"
            )
        mapping[source_key] = target

    if not mapping:
        warnings.append("「字段映射」为空，将按原列名和内置同义词匹配。")
    return mapping


# ---------------------------------------------------------------------------
# 解析辅助
# ---------------------------------------------------------------------------
_WHITESPACE_RE = re.compile(r"\s+", re.UNICODE)
_FONT_SIZE_ALIASES = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5.0,
}


def _find_column(columns: dict[str, int], candidates: tuple[str, ...], *, required: bool = True) -> int | None:
    for candidate in candidates:
        key = norm_header(candidate)
        if key in columns:
            return columns[key]
    if required:
        return None
    return None


def _text_or_default(raw: Any, default: str, *, allow_disable: bool = False) -> str:
    if is_blank(raw):
        return default
    text = to_text(raw)
    if allow_disable and text.strip().lower() in {"无", "none", "null"}:
        return ""
    return text


def _column_name(raw: Any, default: str) -> str:
    text = _text_or_default(raw, default)
    if not text:
        raise ConfigError(f"列名不能为空，请填写「{default}」或映射后的标准列名。")
    return text


def _parse_decimal(raw: Any, label: str) -> Decimal:
    if is_blank(raw):
        raise ConfigError(f"「{label}」不能为空。")
    try:
        return to_decimal(raw)
    except ValueError as exc:
        raise ConfigError(f"「{label}」不是有效数字：{to_text(raw)!r}。") from exc


def _parse_int(raw: Any, label: str, *, minimum: int, maximum: int | None = None) -> int:
    if is_blank(raw):
        raise ConfigError(f"「{label}」不能为空。")
    text = to_text(raw)
    try:
        number = Decimal(text)
    except Exception as exc:  # noqa: BLE001 - 需要统一成中文配置错误
        raise ConfigError(f"「{label}」必须是整数，当前值：{text!r}。") from exc
    if number != number.to_integral_value():
        raise ConfigError(f"「{label}」必须是整数，当前值：{text!r}。")
    value = int(number)
    if value < minimum:
        raise ConfigError(f"「{label}」不能小于 {minimum}，当前值：{value}。")
    if maximum is not None and value > maximum:
        raise ConfigError(f"「{label}」不能大于 {maximum}，当前值：{value}。")
    return value


def _parse_rows_per_page(raw: Any) -> int | None:
    if is_blank(raw):
        raise ConfigError(
            f"「{KEY_ROWS_PER_PAGE}」不能为空。"
            "请填写正整数（例如 18），或填写 None 表示按 Word 自动分页。"
        )
    text = to_text(raw)
    if text.strip().lower() in {"none", "null", "自动", "auto"}:
        return None
    return _parse_int(raw, KEY_ROWS_PER_PAGE, minimum=1)


def _parse_decimals(raw: Any) -> int:
    if is_blank(raw):
        return DEFAULT_DECIMALS
    return _parse_int(raw, KEY_DECIMALS, minimum=0, maximum=4)


def _parse_header_rows(raw: Any) -> int:
    if is_blank(raw):
        return DEFAULT_HEADER_ROWS
    return _parse_int(raw, KEY_HEADER_ROWS, minimum=1, maximum=30)


def _parse_table_index(raw: Any) -> int | None:
    if is_blank(raw):
        return None
    text = to_text(raw)
    if text.strip().lower() in {"自动", "auto", "none", "null"}:
        return None
    return _parse_int(raw, KEY_TABLE_INDEX, minimum=1)


def _parse_font_name(raw: Any) -> str:
    """字体名；空值或“自动”表示沿用原表格识别结果。"""
    if is_blank(raw):
        return DEFAULT_FONT_AUTO
    text = to_text(raw).strip()
    if text.lower() in {"自动", "auto", "none", "null"}:
        return DEFAULT_FONT_AUTO
    return text


def _parse_font_size(raw: Any) -> float | None:
    """字号（磅）；支持 12 / 10.5 / 12pt / 小四；“自动”返回 None。"""
    if is_blank(raw):
        return None
    text = to_text(raw).strip()
    if text.lower() in {"自动", "auto", "none", "null"}:
        return None

    alias = _FONT_SIZE_ALIASES.get(text)
    if alias is not None:
        return alias

    normalized = re.sub(r"(?i)\s*(pt|磅)$", "", text).strip()
    try:
        number = Decimal(normalized)
    except Exception as exc:  # noqa: BLE001
        raise ConfigError(
            f"「{KEY_FONT_SIZE}」不是有效字号：{text!r}。"
            "请填写磅值（如 12、10.5），或“自动”。"
        ) from exc
    if number < 5 or number > 72:
        raise ConfigError(
            f"「{KEY_FONT_SIZE}」必须在 5-72 磅之间，当前值：{text!r}。"
        )
    return float(number)


def _parse_output_pattern(raw: Any, warnings: list[str]) -> str:
    pattern = _text_or_default(raw, DEFAULT_OUTPUT_PATTERN)
    if not pattern:
        pattern = DEFAULT_OUTPUT_PATTERN
    if Path(pattern).name != pattern:
        raise ConfigError(
            f"「{KEY_OUTPUT_NAME}」只能填写文件名，不能包含路径：{pattern!r}。"
        )
    if not pattern.lower().endswith(".docx"):
        warnings.append(f"输出文件名「{pattern}」没有 .docx 后缀，已自动补全。")
        pattern += ".docx"

    try:
        parts = list(string.Formatter().parse(pattern))
    except ValueError as exc:
        raise ConfigError(f"「{KEY_OUTPUT_NAME}」的花括号不匹配：{pattern!r}。") from exc
    for _, field_name, format_spec, conversion in parts:
        if field_name is None:
            continue
        if field_name != "date" or format_spec or conversion:
            raise ConfigError(
                f"「{KEY_OUTPUT_NAME}」只支持占位符 {{date}}，不支持：{pattern!r}。"
            )
    return pattern


def _parse_item_format(format_text: str) -> str:
    try:
        parts = list(string.Formatter().parse(format_text))
    except ValueError as exc:
        raise ConfigError(f"「{KEY_ITEM_FORMAT}」的花括号不匹配：{format_text!r}。") from exc

    allowed = {"name", "amount", "qty", "price"}
    for _, field_name, format_spec, conversion in parts:
        if field_name is None:
            continue
        if conversion or format_spec:
            raise ConfigError(
                f"「{KEY_ITEM_FORMAT}」不支持格式说明或转换符：{format_text!r}。"
                "请直接使用 {name}、{amount}、{qty}、{price}。"
            )
        if field_name not in allowed:
            raise ConfigError(
                f"「{KEY_ITEM_FORMAT}」包含未知占位符 {{{field_name}}}。"
                "可用：{name}、{amount}、{qty}、{price}。"
            )
    try:
        format_text.format(name="值班补贴", amount="22.00", qty="1", price="22.00")
    except Exception as exc:  # noqa: BLE001
        raise ConfigError(f"「{KEY_ITEM_FORMAT}」无法正常使用：{exc}") from exc
    return format_text
