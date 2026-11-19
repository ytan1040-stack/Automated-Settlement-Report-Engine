# -*- coding: utf-8 -*-
"""把 原版表格.docx 转换成 docxtpl 可用的 _内部模板.docx。"""

from __future__ import annotations

import os
import re
from collections import Counter
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table, _Row

from .constants import (
    FIELD_DETAIL,
    FIELD_REMARK,
    MANIFEST_VERSION,
    NON_TOTAL_FIELDS,
)
from .errors import FileLockedError, TemplateError
from .utils import (
    canonical_field,
    norm_header,
    read_json,
    replace_file,
    safe_unlink,
    sha256_file,
    to_text,
    write_json_atomic,
)


@dataclass
class TemplateInfo:
    path: Path
    manifest_path: Path
    manifest: dict[str, Any]
    warnings: list[str] = field(default_factory=list)


@dataclass
class TemplateCheck:
    """模板契约检查结果。

    errors 非空时不能生成模板；warnings 是可以继续、但建议修正的问题。
    """

    source: Path
    table_count: int = 0
    table_index: int = 0
    rows: int = 0
    columns: int = 0
    header_rows: int = 1
    data_row_index: int = 1
    headers: list[str] = field(default_factory=list)
    footer_row: int | None = None
    footer_label: str = ""
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors

    def summary_lines(self) -> list[str]:
        lines = [
            f"主表格：第 {self.table_index or '?'} 个 / 共 {self.table_count} 个",
            f"表格尺寸：{self.rows} 行 × {self.columns} 列",
            f"表头行数：{self.header_rows}；数据行：第 {self.data_row_index + 1} 行",
            "数据列：" + (" | ".join(self.headers) if self.headers else "（无）"),
        ]
        if self.footer_row is None:
            lines.append("表尾合计：未找到")
        else:
            lines.append(
                f"表尾合计：第 {self.footer_row + 1} 行"
                + (f"（{self.footer_label}）" if self.footer_label else "")
            )
        return lines


_NUMBER_RE = re.compile(r"-?\d[\d,]*(?:\.\d+)?")
_TOTAL_KEYWORDS = ("合计", "总计", "小计", "共计")


def check_template(
    source: Path,
    *,
    table_index: int | None = None,
    header_rows: int = 1,
) -> TemplateCheck:
    """只检查原版表格是否符合模板契约，不写出任何文件。

    这是 ``make_templates.py --check`` 的底层实现。检查项包括：
    文件可打开、存在表格、主表可定位、表头/数据行结构、浮动表格、
    嵌套表格、纵向合并、横向合并、表头重复、表尾合计行。
    """
    source = Path(source)
    if not source.exists():
        return TemplateCheck(source=source, errors=[f"未找到「{source.name}」。"])
    if header_rows < 1:
        return TemplateCheck(
            source=source,
            header_rows=header_rows,
            errors=["「表头行数」至少为 1。"],
        )
    try:
        document = Document(str(source))
    except PermissionError as exc:
        raise FileLockedError(
            f"无法读取「{source.name}」：文件可能正在 Word / WPS 中打开。请关闭后重试。"
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise TemplateError(
            f"无法打开「{source.name}」：{exc}。"
            "请确认它是有效的 .docx 文件（旧版 .doc 请先另存为 .docx）。"
        ) from exc

    report, _ = _analyze_template(document, source, table_index, header_rows)
    return report


def _analyze_template(
    document: Any,
    source: Path,
    table_index: int | None,
    header_rows: int,
) -> tuple[TemplateCheck, Table | None]:
    """检查文档并返回诊断结果与选中的主表。"""
    warnings: list[str] = []
    errors: list[str] = []
    report = TemplateCheck(source=source, header_rows=header_rows)

    tables = document.tables
    report.table_count = len(tables)
    if not tables:
        report.errors = ["文档中没有找到任何表格。请确认文档包含 Word 表格。"]
        return report, None

    try:
        chosen_index, table = _pick_table(document, table_index, warnings)
    except TemplateError as exc:
        report.errors = [str(exc)]
        return report, None

    report.table_index = chosen_index
    report.rows = len(table.rows)
    report.columns = len(table.rows[0].cells) if report.rows else 0
    report.data_row_index = header_rows

    tbl_pr = table._tbl.tblPr
    if tbl_pr is not None and tbl_pr.find(qn("w:tblpPr")) is not None:
        errors.append(
            "主表格是浮动表格（文字环绕），自动分页会错位。"
            "请在 Word 中右键表格 →「表格属性」→「文字环绕」选择「无」，保存后重试。"
        )

    if report.rows <= header_rows:
        errors.append(
            f"主表格只有 {report.rows} 行，但「表头行数」为 {header_rows}，找不到数据行。"
            "请调小表头行数，或在表格中保留一行示例数据。"
        )

    data_row = table.rows[header_rows] if report.rows > header_rows else None
    if data_row is not None:
        if _row_has_vertical_merge(data_row):
            errors.append(
                "数据行包含纵向合并单元格，无法自动生成模板。"
                "请先在 Word 中取消数据行的纵向合并，再运行一次。"
            )
        unique = _unique_cells(data_row)
        if not unique:
            errors.append("数据行没有可用的单元格，无法生成模板。")
        elif len(unique) < len(data_row.cells):
            warnings.append(
                "数据行包含横向合并单元格，已按合并后的列生成模板；"
                "如果列对应关系不正确，请先取消数据行的合并单元格。"
            )

    # 嵌套表格会让“一行一个数据对象”的假设失效，直接拒绝。
    nested_error = False
    for row in table.rows:
        for _, cell in _unique_cells(row):
            if cell.tables:
                nested_error = True
                break
        if nested_error:
            break
    if nested_error:
        errors.append(
            "主表格中包含嵌套表格，无法确定数据行边界。"
            "请把子表格移到主表之外，或先把它转成普通单元格内容。"
        )

    if data_row is not None:
        report.headers = []
        for grid_index, _ in _unique_cells(data_row):
            label = _header_label(table, header_rows, grid_index)
            if not label:
                label = f"列{grid_index + 1}"
                warnings.append(f"第 {grid_index + 1} 列没有表头，已命名为「{label}」。")
            report.headers.append(label)

        duplicate_headers = _duplicates(report.headers)
        if duplicate_headers:
            warnings.append(
                "模板中存在重复表头："
                + "、".join(duplicate_headers)
                + "。重复表头会映射到同一数据列，建议改成不同的列名。"
            )

        footer_start = _find_footer_start(table, header_rows + 1)
        report.footer_row = footer_start
        if footer_start is None:
            warnings.append(
                "数据行之后没有找到「合计/总计/小计」行；"
                "生成模板时会删除数据行之后的示例行。"
            )
        else:
            footer_text = " ".join(
                " ".join(cell.text.split()) for cell in table.rows[footer_start].cells
            ).strip()
            report.footer_label = footer_text[:80]
            if not _NUMBER_RE.search(footer_text):
                warnings.append(
                    "表尾合计行中没有数字，可能无法自动求和各列；"
                    "请保留原来的合计数字，程序会用占位符替换。"
                )

    report.errors = errors
    report.warnings = warnings
    return report, table


def build_template(
    source: Path,
    target: Path,
    manifest_path: Path,
    *,
    table_index: int | None = None,
    header_rows: int = 1,
) -> TemplateInfo:
    """读取原版表格，生成带 docxtpl 标签的内部模板与结构清单。"""
    source = Path(source)
    target = Path(target)
    manifest_path = Path(manifest_path)
    if not source.exists():
        raise TemplateError(f"未找到「{source.name}」。请把原始 Word 表格放到程序同级目录。")
    if header_rows < 1:
        raise TemplateError("「表头行数」至少为 1。请在「规则配置.xlsx → 基础配置」中检查。")

    try:
        document = Document(str(source))
    except PermissionError as exc:
        raise FileLockedError(
            f"无法读取「{source.name}」：文件可能正在 Word / WPS 中打开。请关闭后重试。"
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise TemplateError(
            f"无法打开「{source.name}」：{exc}。"
            "请确认它是有效的 .docx 文件（旧版 .doc 请先另存为 .docx）。"
        ) from exc

    report, table = _analyze_template(document, source, table_index, header_rows)
    if table is None or report.errors:
        raise TemplateError(
            "模板不符合生成条件：\n- " + "\n- ".join(report.errors or ["未知错误"])
        )

    chosen_index = report.table_index
    warnings = list(report.warnings)
    data_row_index = report.data_row_index
    data_row = table.rows[data_row_index]

    header_font: dict[str, Any] = {}
    for index in range(header_rows - 1, -1, -1):
        header_font = _capture_font_profile(table.rows[index])
        if header_font:
            break
    data_font = _capture_font_profile(data_row)

    unique_slots = _unique_cells(data_row)
    if not unique_slots:
        raise TemplateError("数据行没有可用的单元格，无法生成模板。")

    footer_start = report.footer_row
    if footer_start is None:
        _delete_rows(table, data_row_index + 1, len(table.rows) - 1)
        footer_rows = 0
    else:
        if footer_start > data_row_index + 1:
            _delete_rows(table, data_row_index + 1, footer_start - 1)
        footer_rows = len(table.rows) - data_row_index - 1

    data_row = table.rows[data_row_index]
    unique_slots = _unique_cells(data_row)
    grid_to_slot: dict[int, int] = {}
    starts = [grid for grid, _ in unique_slots]
    total_grid_columns = len(data_row.cells)

    headers = list(report.headers)
    slots: list[str] = []
    for order, (grid_index, cell) in enumerate(unique_slots, start=1):
        header = headers[order - 1] if order <= len(headers) else f"列{grid_index + 1}"
        slots.append(f"col{order}")
        _set_cell_text(cell, f"{{{{ item.col{order} }}}}")

        end = starts[order] if order < len(starts) else total_grid_columns
        for grid in range(grid_index, max(grid_index + 1, end)):
            grid_to_slot.setdefault(grid, order)

    total_slots = _prepare_footer_totals(
        table,
        header_rows,
        data_row_index,
        grid_to_slot,
        warnings,
    )

    for row_index in range(header_rows):
        _set_row_flag(table.rows[row_index]._tr, "tblHeader")
    _set_row_flag(data_row._tr, "cantSplit")

    _insert_loop_markers(table, data_row_index, data_row)

    manifest: dict[str, Any] = {
        "version": MANIFEST_VERSION,
        "source": source.name,
        "source_sha256": sha256_file(source),
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "table_index": chosen_index,
        "header_rows": header_rows,
        "data_row_index": data_row_index,
        "footer_rows": footer_rows,
        "headers": headers,
        "slots": slots,
        "loop_var": "item",
        "loop_list": "rows",
        "total_slots": total_slots,
        "fonts": {
            "header": header_font,
            "data": data_font,
        },
    }

    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_name(f".{target.stem}.{os.getpid()}.tmp.docx")
    try:
        document.save(str(temp))
        replace_file(temp, target)
    finally:
        safe_unlink(temp)
    write_json_atomic(manifest_path, manifest)

    return TemplateInfo(
        path=target,
        manifest_path=manifest_path,
        manifest=manifest,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# 表格选择
# ---------------------------------------------------------------------------
def _pick_table(document: Any, table_index: int | None, warnings: list[str]) -> tuple[int, Table]:
    tables = document.tables
    if not tables:
        raise TemplateError("「原版表格.docx」中没有找到任何表格。请确认文档包含 Word 表格。")
    if table_index is not None:
        if table_index < 1 or table_index > len(tables):
            raise TemplateError(
                f"「主表格序号」为 {table_index}，但文档中只有 {len(tables)} 个表格。"
                "请修改「规则配置.xlsx → 基础配置 → 主表格序号」。"
            )
        return table_index, tables[table_index - 1]

    def score(table: Table) -> tuple[int, int, int]:
        rows = len(table.rows)
        columns = len(table.rows[0].cells) if rows else 0
        return rows * columns, rows, columns

    best = max(range(len(tables)), key=lambda index: score(tables[index]))
    if len(tables) > 1:
        best_rows = len(tables[best].rows)
        best_columns = len(tables[best].rows[0].cells) if best_rows else 0
        warnings.append(
            f"文档中有 {len(tables)} 个表格，已自动选择第 {best + 1} 个"
            f"（{best_rows} 行 × {best_columns} 列）。"
            "如需指定，请修改「主表格序号」。"
        )
    return best + 1, tables[best]


# ---------------------------------------------------------------------------
# 行/列辅助
# ---------------------------------------------------------------------------
def _unique_cells(row: _Row) -> list[tuple[int, Any]]:
    seen: set[int] = set()
    result: list[tuple[int, Any]] = []
    for grid_index, cell in enumerate(row.cells):
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        result.append((grid_index, cell))
    return result


def _header_label(table: Table, header_rows: int, grid_index: int) -> str:
    for row_index in range(header_rows - 1, -1, -1):
        cells = table.rows[row_index].cells
        if grid_index >= len(cells):
            continue
        text = " ".join(cells[grid_index].text.split())
        if text:
            return text
    return ""


def _find_footer_start(table: Table, start: int) -> int | None:
    for row_index in range(start, len(table.rows)):
        if _is_total_row(table.rows[row_index]):
            return row_index
    return None


def _is_total_row(row: _Row) -> bool:
    for cell in row.cells:
        text = norm_header(cell.text)
        if any(keyword in text for keyword in _TOTAL_KEYWORDS):
            return True
    return False


def _delete_rows(table: Table, start: int, end: int) -> None:
    if end < start:
        return
    for row_index in range(end, start - 1, -1):
        table._tbl.remove(table.rows[row_index]._tr)


def _row_has_vertical_merge(row: _Row) -> bool:
    for cell in _unique_cells(row):
        tc_pr = cell[1]._tc.tcPr
        if tc_pr is not None and tc_pr.find(qn("w:vMerge")) is not None:
            return True
    return False


def _ensure_no_vertical_merge(row: _Row) -> None:
    if _row_has_vertical_merge(row):
        raise TemplateError(
            "数据行包含纵向合并单元格，无法自动生成模板。"
            "请先在 Word 中取消数据行的纵向合并，再运行一次。"
        )


# ---------------------------------------------------------------------------
# 单元格写入
# ---------------------------------------------------------------------------
def _capture_font_profile(row: _Row) -> dict[str, Any]:
    """从一行单元格中提取出现最多的中文/西文字体与字号。

    Word 的字体既可能直接写在 run 上，也可能继承自段落样式；这里两种都读。
    取“出现最多”的值，避免某一两个特殊 run（如数字、符号）带偏整行。
    """
    east_names: Counter[str] = Counter()
    ascii_names: Counter[str] = Counter()
    sizes: Counter[float] = Counter()

    for _, cell in _unique_cells(row):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                east_name, ascii_name, size = _run_font_values(run, paragraph)
                if east_name:
                    east_names[east_name] += 1
                if ascii_name:
                    ascii_names[ascii_name] += 1
                if size:
                    sizes[round(size, 2)] += 1

    profile: dict[str, Any] = {}
    if east_names:
        profile["east_asia"] = east_names.most_common(1)[0][0]
    if ascii_names:
        profile["ascii"] = ascii_names.most_common(1)[0][0]
    if sizes:
        profile["size_pt"] = sizes.most_common(1)[0][0]
    return profile


def _run_font_values(run: Any, paragraph: Any) -> tuple[str | None, str | None, float | None]:
    east_name: str | None = None
    ascii_name: str | None = None
    size = run.font.size.pt if run.font.size is not None else None

    r_pr = run._element.rPr
    if r_pr is not None and r_pr.rFonts is not None:
        r_fonts = r_pr.rFonts
        east_name = r_fonts.get(qn("w:eastAsia")) or None
        ascii_name = (
            r_fonts.get(qn("w:ascii"))
            or r_fonts.get(qn("w:hAnsi"))
            or None
        )

    if not ascii_name:
        ascii_name = run.font.name

    style = paragraph.style
    if style is not None:
        style_font = style.font
        if not ascii_name:
            ascii_name = style_font.name
        if size is None and style_font.size is not None:
            size = style_font.size.pt
        if not east_name:
            try:
                style_r_pr = style.element.rPr
            except AttributeError:
                style_r_pr = None
            if style_r_pr is not None and style_r_pr.rFonts is not None:
                east_name = style_r_pr.rFonts.get(qn("w:eastAsia")) or None

    return east_name, ascii_name, size


def _set_cell_text(cell: Any, text: str) -> None:
    paragraphs = cell.paragraphs
    if not paragraphs:
        cell.text = text
        return

    paragraph = paragraphs[0]
    run_properties = None
    for run in paragraph.runs:
        if run._element.rPr is not None:
            run_properties = deepcopy(run._element.rPr)
            break

    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)

    run = paragraph.add_run(text)
    if run_properties is not None:
        run._element.insert(0, run_properties)

    for extra in paragraphs[1:]:
        cell._tc.remove(extra._p)


def _set_row_flag(tr: Any, tag: str) -> None:
    tr_pr = tr.get_or_add_trPr()
    if tr_pr.find(qn(f"w:{tag}")) is None:
        tr_pr.append(OxmlElement(f"w:{tag}"))


# ---------------------------------------------------------------------------
# 循环标记
# ---------------------------------------------------------------------------
def _insert_loop_markers(table: Table, data_row_index: int, data_row: _Row) -> None:
    for_tag = "{%tr for item in rows %}"
    end_tag = "{%tr endfor %}"

    for_tr = deepcopy(data_row._tr)
    data_row._tr.addprevious(for_tr)
    for_row = _Row(for_tr, table)
    _fill_marker_row(for_row, for_tag)

    end_tr = deepcopy(data_row._tr)
    data_row._tr.addnext(end_tr)
    end_row = _Row(end_tr, table)
    _fill_marker_row(end_row, end_tag)


def _fill_marker_row(row: _Row, tag: str) -> None:
    for cell in row.cells:
        _set_cell_text(cell, "")
    unique = _unique_cells(row)
    if unique:
        _set_cell_text(unique[0][1], tag)


# ---------------------------------------------------------------------------
# 表尾合计
# ---------------------------------------------------------------------------
def _prepare_footer_totals(
    table: Table,
    header_rows: int,
    data_row_index: int,
    grid_to_slot: dict[int, int],
    warnings: list[str],
) -> list[dict[str, Any]]:
    total_slots: list[dict[str, Any]] = []
    seen_slots: set[int] = set()

    for row_index in range(data_row_index + 1, len(table.rows)):
        row = table.rows[row_index]
        if not _is_total_row(row):
            continue
        for grid_index, cell in _unique_cells(row):
            slot = grid_to_slot.get(grid_index)
            if slot is None:
                continue
            header = _header_label(table, header_rows, grid_index)
            canonical = canonical_field(header)
            if not canonical:
                continue
            if canonical in NON_TOTAL_FIELDS or canonical in {FIELD_DETAIL, FIELD_REMARK}:
                continue
            if any(keyword in canonical for keyword in ("日期", "时间", "身份证", "手机", "电话")):
                continue
            if slot in seen_slots:
                continue

            placeholder = f"total_col{slot}"
            current = " ".join(cell.text.split())
            if "{{" in current:
                continue
            replaced = False
            if not current:
                _set_cell_text(cell, f"{{{{ {placeholder} }}}}")
                replaced = True
            else:
                matches = list(_NUMBER_RE.finditer(current))
                if matches:
                    match = matches[-1]
                    _set_cell_text(
                        cell,
                        current[: match.start()] + f"{{{{ {placeholder} }}}}" + current[match.end() :],
                    )
                    replaced = True
                else:
                    replaced = False
            if replaced:
                seen_slots.add(slot)
                total_slots.append(
                    {
                        "slot": slot,
                        "header": header,
                        "placeholder": placeholder,
                    }
                )

    if not total_slots:
        warnings.append("表尾没有可自动求和的数字单元格；如果模板有合计行，请确认列名是否匹配。")
    return total_slots


# ---------------------------------------------------------------------------
# 清单
# ---------------------------------------------------------------------------
def load_manifest(path: Path) -> dict[str, Any]:
    path = Path(path)
    if not path.exists():
        raise TemplateError(
            f"未找到模板结构文件「{path.name}」。"
            "请先运行 make_templates.py，或直接双击 一键运行.bat。"
        )
    try:
        manifest = read_json(path)
    except Exception as exc:  # noqa: BLE001
        raise TemplateError(f"无法读取「{path.name}」：{exc}。请重新运行 make_templates.py。") from exc

    if not isinstance(manifest, dict):
        raise TemplateError(f"「{path.name}」格式不正确，请重新运行 make_templates.py。")
    if manifest.get("version") != MANIFEST_VERSION:
        raise TemplateError(
            f"「{path.name}」版本过旧或不兼容，请重新运行 make_templates.py 生成模板。"
        )
    for key in ("table_index", "header_rows", "footer_rows", "headers", "slots"):
        if key not in manifest:
            raise TemplateError(
                f"「{path.name}」缺少字段「{key}」，请重新运行 make_templates.py。"
            )
    if len(manifest["headers"]) != len(manifest["slots"]):
        raise TemplateError(f"「{path.name}」的表头与占位符数量不一致，请重新生成模板。")
    return manifest


def _duplicates(items: list[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for item in items:
        key = to_text(item)
        if key in seen and key not in duplicates:
            duplicates.append(key)
        seen.add(key)
    return duplicates
