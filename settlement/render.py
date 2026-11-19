# -*- coding: utf-8 -*-
"""docxtpl 渲染与分页后处理。"""

from __future__ import annotations

import os
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined, TemplateSyntaxError, UndefinedError

from .errors import FileLockedError, ReportError
from .utils import replace_file, safe_unlink


@dataclass
class RenderResult:
    path: Path
    data_rows: int
    page_breaks: int
    warnings: list[str] = field(default_factory=list)


def render_report(
    template_path: Path,
    output_path: Path,
    context: dict[str, Any],
    *,
    manifest: dict[str, Any],
    rows_per_page: int | None,
    expected_rows: int,
    font_east_asia: str = "自动",
    font_ascii: str = "自动",
    font_size_pt: float | None = None,
) -> RenderResult:
    """渲染模板、按每页行数拆分表格分页，并原子写出最终文件。"""
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_render = output_path.with_name(f".{output_path.stem}.{os.getpid()}.render.docx")
    temp_final = output_path.with_name(f".{output_path.stem}.{os.getpid()}.final.docx")
    try:
        try:
            template = DocxTemplate(str(template_path))
            environment = Environment(undefined=StrictUndefined, autoescape=False)
            template.render(context, jinja_env=environment)
            template.save(str(temp_render))
        except UndefinedError as exc:
            raise ReportError(
                f"模板中有无法解析的占位符：{exc}。"
                "请确认模板列名与源数据列名能对应，或重新运行 make_templates.py。"
            ) from exc
        except TemplateSyntaxError as exc:
            raise ReportError(f"模板标签语法错误：{exc}。请重新生成 _内部模板.docx。") from exc
        except PermissionError as exc:
            raise FileLockedError(
                f"无法写入「{output_path.name}」：文件可能正在 Word / WPS 中打开。请关闭后重试。"
            ) from exc
        except Exception as exc:  # noqa: BLE001
            raise ReportError(f"渲染 Word 失败：{exc}") from exc

        try:
            document = Document(str(temp_render))
        except Exception as exc:  # noqa: BLE001
            raise ReportError(f"无法读取渲染后的 Word：{exc}") from exc

        warnings: list[str] = []
        _apply_fonts(
            document,
            manifest,
            font_east_asia,
            font_ascii,
            font_size_pt,
        )
        data_rows, page_breaks = _apply_pagination(
            document,
            manifest,
            rows_per_page,
            expected_rows,
            warnings,
        )
        document.save(str(temp_final))
        replace_file(temp_final, output_path)
    finally:
        safe_unlink(temp_render)
        safe_unlink(temp_final)

    return RenderResult(
        path=output_path,
        data_rows=data_rows,
        page_breaks=page_breaks,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# 字体：自动沿用模板识别结果，或按配置覆盖
# ---------------------------------------------------------------------------
def _apply_fonts(
    document: Any,
    manifest: dict[str, Any],
    font_east_asia: str,
    font_ascii: str,
    font_size_pt: float | None,
) -> None:
    table_index = int(manifest.get("table_index", 1))
    if table_index < 1 or table_index > len(document.tables):
        raise ReportError(
            f"渲染后的文档中找不到第 {table_index} 个表格，模板可能已被修改。"
            "请重新运行 make_templates.py。"
        )

    table = document.tables[table_index - 1]
    header_rows = int(manifest.get("header_rows", 1))
    fonts = manifest.get("fonts") or {}
    header_profile = fonts.get("header") or {}
    data_profile = fonts.get("data") or {}

    requested_east = _requested_font(font_east_asia)
    requested_ascii = _requested_font(font_ascii)
    requested_size = float(font_size_pt) if font_size_pt is not None else None

    for row_index, row in enumerate(table.rows):
        profile = header_profile if row_index < header_rows else data_profile
        east = requested_east or profile.get("east_asia")
        ascii_name = requested_ascii or profile.get("ascii")
        size = requested_size if requested_size is not None else profile.get("size_pt")
        _apply_row_fonts(row, east, ascii_name, size)


def _requested_font(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() in {"自动", "auto", "none", "null"}:
        return None
    return text


def _apply_row_fonts(
    row: Any,
    east_asia: Any,
    ascii_name: Any,
    size_pt: Any,
) -> None:
    east_text = str(east_asia).strip() if east_asia else None
    ascii_text = str(ascii_name).strip() if ascii_name else None
    size = float(size_pt) if size_pt else None

    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if ascii_text:
                    run.font.name = ascii_text
                if east_text:
                    r_pr = run._element.get_or_add_rPr()
                    r_fonts = r_pr.get_or_add_rFonts()
                    r_fonts.set(qn("w:eastAsia"), east_text)
                if size is not None:
                    run.font.size = Pt(size)


# ---------------------------------------------------------------------------
# 分页：把渲染后的单张大表按每页行数拆成多张表
# ---------------------------------------------------------------------------
def _apply_pagination(
    document: Any,
    manifest: dict[str, Any],
    rows_per_page: int | None,
    expected_rows: int,
    warnings: list[str],
) -> tuple[int, int]:
    table_index = int(manifest.get("table_index", 1))
    if table_index < 1 or table_index > len(document.tables):
        raise ReportError(
            f"渲染后的文档中找不到第 {table_index} 个表格，模板可能已被修改。"
            "请重新运行 make_templates.py。"
        )
    table = document.tables[table_index - 1]
    data_start = int(manifest.get("header_rows", 1))
    footer_rows = int(manifest.get("footer_rows", 0))
    available = len(table.rows) - data_start - footer_rows
    if available < 0:
        raise ReportError("模板行数与结构清单不一致，请重新运行 make_templates.py 生成模板。")
    if available != expected_rows:
        warnings.append(
            f"渲染后数据行数为 {available}，与源数据行数 {expected_rows} 不一致，"
            "分页位置可能不准确，请检查模板。"
        )

    for index in range(data_start):
        _set_row_flag(table.rows[index]._tr, "tblHeader")
    for index in range(data_start, data_start + available + footer_rows):
        _set_row_flag(table.rows[index]._tr, "cantSplit")

    if not rows_per_page or available <= rows_per_page:
        return available, 0

    chunks = [
        list(range(start, min(start + rows_per_page, available)))
        for start in range(0, available, rows_per_page)
    ]
    _split_table(table, data_start, footer_rows, chunks)
    return available, len(chunks) - 1


def _split_table(
    table: Any,
    data_start: int,
    footer_rows: int,
    chunks: list[list[int]],
) -> None:
    """把一张表按数据行分块，每块一张表，块间插入分页符。"""
    tbl = table._tbl
    body = tbl.getparent()
    source_rows = tbl.findall(qn("w:tr"))
    total_rows = len(source_rows)
    body_rows = data_start + sum(len(chunk) for chunk in chunks) + footer_rows
    if total_rows != body_rows:
        raise ReportError("分页时表格行数不一致，请重新运行 make_templates.py。")

    elements: list[Any] = []
    for page_index, chunk in enumerate(chunks):
        clone = deepcopy(tbl)
        clone_rows = clone.findall(qn("w:tr"))
        keep = set(range(data_start)) | {data_start + index for index in chunk}
        if page_index == len(chunks) - 1 and footer_rows:
            keep.update(range(data_start + sum(len(item) for item in chunks), total_rows))
        for row_index in range(len(clone_rows) - 1, -1, -1):
            if row_index not in keep:
                clone.remove(clone_rows[row_index])

        _decorate_chunk(clone, data_start, len(chunk), footer_rows if page_index == len(chunks) - 1 else 0)
        elements.append(clone)
        if page_index < len(chunks) - 1:
            elements.append(_page_break_paragraph())

    for element in elements:
        tbl.addprevious(element)
    body.remove(tbl)


def _decorate_chunk(clone: Any, data_start: int, data_count: int, footer_rows: int) -> None:
    rows = clone.findall(qn("w:tr"))
    for index in range(min(data_start, len(rows))):
        _set_row_flag(rows[index], "tblHeader")
    for index in range(data_start, min(data_start + data_count, len(rows))):
        _set_row_flag(rows[index], "cantSplit")
    for index in range(len(rows) - footer_rows, len(rows)):
        _set_row_flag(rows[index], "cantSplit")


def _set_row_flag(tr: Any, flag: str) -> None:
    tr_pr = tr.get_or_add_trPr()
    if tr_pr.find(qn(f"w:{flag}")) is None:
        tr_pr.append(OxmlElement(f"w:{flag}"))


def _page_break_paragraph() -> Any:
    """一个只有分页符、高度约 1 磅的空段落，用来分隔两张表。"""
    paragraph = OxmlElement("w:p")
    paragraph_properties = OxmlElement("w:pPr")

    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")
    paragraph_properties.append(spacing)

    run_properties = OxmlElement("w:rPr")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "2")
    run_properties.append(size)
    paragraph_properties.append(run_properties)
    paragraph.append(paragraph_properties)

    run = OxmlElement("w:r")
    run.append(deepcopy(run_properties))
    page_break = OxmlElement("w:br")
    page_break.set(qn("w:type"), "page")
    run.append(page_break)
    paragraph.append(run)
    return paragraph
