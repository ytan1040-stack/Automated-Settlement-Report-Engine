# -*- coding: utf-8 -*-
"""Excel 读写辅助：统一处理文件占用、损坏文件与表头定位。"""

from __future__ import annotations

import os
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .errors import ConfigError, DataError, FileLockedError, SettlementError
from .utils import norm_header, replace_file, safe_unlink, to_text


def load_workbook_safe(path: Path, *, data_only: bool = True) -> Workbook:
    """读取 xlsx；把常见的“被占用/不是 xlsx”转换成用户能看懂的异常。"""
    if not path.exists():
        raise SettlementError(f"未找到文件：{path.name}")
    try:
        return load_workbook(filename=str(path), data_only=data_only)
    except PermissionError as exc:
        raise FileLockedError(
            f"无法读取「{path.name}」：文件可能正在 Excel 中打开。请关闭后重试。"
        ) from exc
    except (InvalidFileException, BadZipFile, KeyError) as exc:
        raise SettlementError(
            f"「{path.name}」不是有效的 .xlsx 文件。"
            "如果是旧版 .xls，请用 Excel 另存为 .xlsx 后重试。"
        ) from exc
    except OSError as exc:
        raise SettlementError(f"读取「{path.name}」失败：{exc}") from exc


def save_workbook_atomic(workbook: Workbook, path: Path) -> None:
    """先写临时文件再原子替换，避免生成一半的损坏文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.stem}.{os.getpid()}.tmp.xlsx")
    try:
        workbook.save(str(temp))
        replace_file(temp, path)
    finally:
        safe_unlink(temp)


def find_sheet(workbook: Workbook, preferred_name: str, index: int) -> Worksheet:
    """优先按名称找 Sheet；用户重命名后按位置兜底。"""
    if preferred_name in workbook.sheetnames:
        return workbook[preferred_name]
    if 0 <= index < len(workbook.worksheets):
        return workbook.worksheets[index]
    raise ConfigError(
        f"规则配置缺少工作表「{preferred_name}」，且第 {index + 1} 个工作表不存在。"
        "请重新运行 init_config.py 生成配置模板。"
    )


def find_header_row(
    worksheet: Worksheet,
    candidates: tuple[str, ...],
    *,
    max_scan: int = 15,
) -> tuple[int | None, dict[str, int]]:
    """在前若干行中查找表头，返回 (行号, 归一化表头 -> 列号)。

    行号从 1 开始。找不到时返回 (None, {})，由调用方决定兜底或报错。
    """
    wanted = {norm_header(item) for item in candidates}
    limit = min(worksheet.max_row, max_scan)
    for row_index in range(1, limit + 1):
        values = [norm_header(cell.value) for cell in worksheet[row_index]]
        if not any(value in wanted for value in values):
            continue
        columns: dict[str, int] = {}
        for column_index, value in enumerate(values, start=1):
            if value and value not in columns:
                columns[value] = column_index
        return row_index, columns
    return None, {}


def header_row_values(worksheet: Worksheet, row_index: int) -> list[str]:
    return [to_text(cell.value) for cell in worksheet[row_index]]


def require_columns(
    columns: dict[str, int],
    required: tuple[str, ...],
    *,
    sheet_name: str,
) -> None:
    missing = [name for name in required if norm_header(name) not in columns]
    if missing:
        raise ConfigError(
            f"「{sheet_name}」缺少必需的表头：{'、'.join(missing)}。"
            "请重新运行 init_config.py，或按模板补齐表头。"
        )


def cell_value(worksheet: Worksheet, row: int, column: int) -> object:
    return worksheet.cell(row=row, column=column).value


def is_empty_row(worksheet: Worksheet, row: int, columns: int) -> bool:
    for column in range(1, columns + 1):
        value = worksheet.cell(row=row, column=column).value
        if not (value is None or (isinstance(value, str) and not value.strip())):
            return False
    return True


def data_error(sheet_name: str, row: int, message: str) -> DataError:
    return DataError(f"「{sheet_name}」第 {row} 行：{message}")
