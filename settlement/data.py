# -*- coding: utf-8 -*-
"""读取 源数据.xlsx：定位表头、应用字段映射、清洗单元格。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from openpyxl.worksheet.worksheet import Worksheet

from .constants import (
    DERIVED_ALIASES,
    FIELD_SEQ,
)
from .errors import DataError
from .excel import find_header_row, load_workbook_safe
from .utils import canonical_field, fmt_number, is_blank, norm_header, to_text


@dataclass
class DataTable:
    """一张已经完成表头映射的数据表。"""

    path: Path
    sheet_name: str
    header_row: int
    headers: tuple[str, ...]
    fields: tuple[str, ...]
    records: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.records)


_DEFAULT_HEADER_CANDIDATES: tuple[str, ...] = tuple(
    dict.fromkeys(
        [FIELD_SEQ, "姓名", "员工姓名", "部门", "工号", "固定基数", "浮动基数", "备注"]
        + [alias for aliases in DERIVED_ALIASES.values() for alias in aliases]
    )
)


def load_data(
    path: Path,
    mapping: dict[str, str] | None = None,
    *,
    sheet_name: str | None = None,
    header_candidates: Iterable[str] | None = None,
) -> DataTable:
    """读取源数据并返回标准字段记录。

    mapping 的 key 是归一化后的原始列名，value 是归一化后的标准字段名。
    header_candidates 用于在复杂表头中定位真正的表头行；找不到时会自动兜底。
    """
    path = Path(path)
    workbook = load_workbook_safe(path, data_only=True)
    try:
        worksheet = _pick_sheet(workbook, sheet_name, path)
        candidates = tuple(
            dict.fromkeys(
                list(header_candidates or ()) + list(_DEFAULT_HEADER_CANDIDATES)
            )
        )
        header_row, _ = find_header_row(worksheet, candidates, max_scan=12)
        warnings: list[str] = []

        if header_row is None:
            header_row = _guess_header_row(worksheet)
            if header_row is None:
                raise DataError(
                    f"无法在「{path.name}」中定位表头。"
                    "请确保数据表第 1 行是列名（可参考「源数据模板.xlsx」）。"
                )
            warnings.append(
                f"未能识别标准表头，已将第 {header_row} 行作为表头，请确认列名是否正确。"
            )

        table = _read_table(worksheet, header_row, mapping or {}, path, warnings)
    finally:
        workbook.close()
    return table


# ---------------------------------------------------------------------------
# 工作表与表头
# ---------------------------------------------------------------------------
def _pick_sheet(workbook: Any, sheet_name: str | None, path: Path) -> Worksheet:
    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            raise DataError(
                f"「{path.name}」中没有名为「{sheet_name}」的工作表。"
                f"现有工作表：{'、'.join(workbook.sheetnames)}。"
            )
        return workbook[sheet_name]
    if workbook.active is not None:
        return workbook.active
    if workbook.worksheets:
        return workbook.worksheets[0]
    raise DataError(f"「{path.name}」中没有任何工作表。")


def _guess_header_row(worksheet: Worksheet, max_scan: int = 10) -> int | None:
    """没有标准列名时，找第一行“至少两个单元格且以文本为主”的行。"""
    limit = min(worksheet.max_row, max_scan)
    for row in range(1, limit + 1):
        values = [cell.value for cell in worksheet[row]]
        non_empty = [value for value in values if not is_blank(value)]
        if len(non_empty) < 2:
            continue
        text_count = sum(1 for value in non_empty if isinstance(value, str))
        if text_count / len(non_empty) >= 0.6:
            return row
    return None


# ---------------------------------------------------------------------------
# 读取正文
# ---------------------------------------------------------------------------
def _read_table(
    worksheet: Worksheet,
    header_row: int,
    mapping: dict[str, str],
    path: Path,
    warnings: list[str],
) -> DataTable:
    max_column = max(1, worksheet.max_column)
    raw_headers = [
        to_text(worksheet.cell(row=header_row, column=column).value)
        for column in range(1, max_column + 1)
    ]

    headers: list[str] = []
    fields: list[str] = []
    column_map: list[tuple[int, str, str]] = []
    used_fields: dict[str, str] = {}
    normalized_headers: set[str] = set()

    for column, raw_header in enumerate(raw_headers, start=1):
        if not raw_header:
            continue
        key = norm_header(raw_header)
        if not key:
            continue
        normalized_headers.add(key)

        target = mapping.get(key)
        if target:
            target = norm_header(target)
        else:
            target = norm_header(canonical_field(raw_header))
        if not target:
            target = key

        if target in used_fields:
            raise DataError(
                f"「{path.name}」第 {header_row} 行：列「{raw_header}」与"
                f"「{used_fields[target]}」映射到了同一个字段「{target}」。"
                "请修改「规则配置.xlsx → 字段映射」，确保一列只对应一个标准字段。"
            )

        used_fields[target] = raw_header
        headers.append(raw_header)
        fields.append(target)
        column_map.append((column, raw_header, target))

    if not fields:
        raise DataError(
            f"「{path.name}」第 {header_row} 行没有可用的列名。"
            "请参考「源数据模板.xlsx」填写表头。"
        )

    _warn_unused_mapping(mapping, normalized_headers, set(used_fields), warnings)

    records: list[dict[str, Any]] = []
    for row in range(header_row + 1, worksheet.max_row + 1):
        record: dict[str, Any] = {}
        has_value = False
        for column, _header, field_name in column_map:
            value = _normalize_cell(worksheet.cell(row=row, column=column).value)
            record[field_name] = value
            if value != "":
                has_value = True
        if not has_value:
            continue
        records.append(record)

    if not records:
        raise DataError(
            f"「{path.name}」第 {header_row} 行以下没有数据。"
            "请填写至少一行数据后重试。"
        )

    if FIELD_SEQ not in used_fields:
        # 模板里如果有“序号”列，仍然可以显示自动序号。
        for index, record in enumerate(records, start=1):
            record.setdefault(FIELD_SEQ, str(index))
    else:
        for index, record in enumerate(records, start=1):
            if is_blank(record.get(FIELD_SEQ)):
                record[FIELD_SEQ] = str(index)

    if len(records) > 5000:
        warnings.append(
            f"共读取 {len(records)} 行数据，生成 Word 可能需要较长时间，请耐心等待。"
        )

    return DataTable(
        path=path,
        sheet_name=worksheet.title,
        header_row=header_row,
        headers=tuple(headers),
        fields=tuple(fields),
        records=records,
        warnings=warnings,
    )


def _normalize_cell(value: Any) -> Any:
    """把 Excel 值变成适合 Word 展示的值；数字去掉多余的 0，日期转成文本。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, (datetime, date, time)):
        return to_text(value)
    if isinstance(value, Decimal):
        return fmt_number(value, 6)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return fmt_number(value, 6)
    if isinstance(value, str):
        return value.strip()
    return to_text(value)


def _warn_unused_mapping(
    mapping: dict[str, str],
    normalized_headers: set[str],
    used_fields: set[str],
    warnings: list[str],
) -> None:
    missing = [
        key
        for key, target in mapping.items()
        if key not in normalized_headers and target not in used_fields
    ]
    for key in missing:
        warnings.append(f"字段映射中的「{key}」在源数据里没有对应列，已忽略。")
