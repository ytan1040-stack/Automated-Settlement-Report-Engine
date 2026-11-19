# -*- coding: utf-8 -*-
"""生成「规则配置.xlsx」与「源数据模板.xlsx」。"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from .constants import (
    CONFIG_FILE,
    DATA_TEMPLATE_FILE,
    FIELD_DEPARTMENT,
    FIELD_EMPLOYEE_ID,
    FIELD_FIXED_BASE,
    FIELD_FLOATING_BASE,
    FIELD_NAME,
    FIELD_REMARK,
    FIELD_SEQ,
    PROJECT_TITLE,
    SHEET_BASE,
    SHEET_ITEMS,
    SHEET_MAPPING,
)
from .excel import save_workbook_atomic
from .utils import info

TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL = PatternFill("solid", fgColor="F2F7FC")
WHITE_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
BODY_FONT = Font(name="宋体", size=11, color="1F1F1F")
NOTE_FONT = Font(name="宋体", size=9, color="7F6000")
THIN = Side(style="thin", color="B4C6E7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def create_config(path: Path, *, overwrite: bool = False) -> Path:
    """写出规则配置工作簿；overwrite=False 时保留已有文件。"""
    path = Path(path)
    if path.exists() and not overwrite:
        info(f"已存在「{path.name}」，保留现有配置")
        return path

    wb = Workbook()
    _build_base_sheet(wb.active)
    _build_items_sheet(wb.create_sheet(SHEET_ITEMS))
    _build_mapping_sheet(wb.create_sheet(SHEET_MAPPING))
    wb.active = 0

    save_workbook_atomic(wb, path)
    info(f"已生成规则配置：{path.name}")
    return path


def create_data_template(path: Path, *, overwrite: bool = False) -> Path:
    """写出源数据模板；overwrite=False 时保留已有文件。"""
    path = Path(path)
    if path.exists() and not overwrite:
        info(f"已存在「{path.name}」，保留现有文件")
        return path

    wb = Workbook()
    _build_data_sheet(wb.active)
    _build_help_sheet(wb.create_sheet("填写说明"))
    wb.active = 0

    save_workbook_atomic(wb, path)
    info(f"已生成源数据模板：{path.name}")
    return path


# ---------------------------------------------------------------------------
# Sheet1 基础配置
# ---------------------------------------------------------------------------
def _build_base_sheet(ws: Worksheet) -> None:
    ws.title = SHEET_BASE
    ws.sheet_properties.tabColor = "1F4E79"
    _title(ws, "基础配置（带 * 为必填；修改后保存，再运行一键脚本）", 4)

    headers = ["配置项", "配置值", "说明", "示例"]
    _headers(ws, 2, headers)

    rows: list[tuple[object, object, str, str]] = [
        ("固定基数单价", 15, "按固定基数计费的单价（元/单位）", "15"),
        ("浮动基数单价", 24, "按浮动基数计费的单价（元/单位）", "24"),
        (
            "单项展示格式",
            "{name}+{amount}",
            "明细中每个计费项的格式；可用 {name}、{amount}、{qty}、{price}",
            "值班+22.00",
        ),
        ("多项分隔符", ",", "多个计费项之间的分隔符", ","),
        ("总计前缀", ", 共+", "明细末尾合计的前缀；留空则只显示合计", ", 共+"),
        (
            "每页固定行数",
            18,
            "正整数；填 None 表示按 Word 自动流式分页",
            "18 或 None",
        ),
        ("固定基数列", FIELD_FIXED_BASE, "源数据中固定基数列的标准列名", FIELD_FIXED_BASE),
        ("浮动基数列", FIELD_FLOATING_BASE, "源数据中浮动基数列的标准列名", FIELD_FLOATING_BASE),
        ("报表标题", "结算报表", "可写入模板 {{ report_title }}", "结算报表"),
        ("金额小数位", 2, "0-4，影响明细与合计金额格式", "2"),
        (
            "输出文件名",
            "结算报表_{date}.docx",
            "可用 {date} 占位符，自动替换为 YYYYMMDD",
            "结算报表_20260101.docx",
        ),
        (
            "主表格序号",
            "自动",
            "文档中有多个表格时填 1、2…；建议保持自动",
            "自动",
        ),
        ("表头行数", 1, "原版表格中表头占几行", "1"),
        (
            "中文字体",
            "自动",
            "自动=沿用原表格识别结果；也可填 宋体、仿宋_GB2312、黑体 等",
            "自动",
        ),
        (
            "西文字体",
            "自动",
            "自动=沿用原表格识别结果；也可填 Times New Roman、Arial 等",
            "自动",
        ),
        (
            "输出字号",
            "自动",
            "自动=沿用原表格识别结果；也可填 12、10.5 或 小四",
            "自动",
        ),
    ]

    for index, (name, value, note, example) in enumerate(rows, start=3):
        ws.cell(index, 1, name)
        ws.cell(index, 2, value)
        ws.cell(index, 3, note)
        ws.cell(index, 4, example)
        for column in range(1, 5):
            cell = ws.cell(index, column)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = LEFT if column in (1, 3, 4) else CENTER
            if index % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[index].height = 24

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 26
    ws.freeze_panes = "A3"

    _whole_validation(ws, "B12", 0, 4, "金额小数位", "只能填写 0-4 的整数")
    _whole_validation(ws, "B15", 0, 30, "表头行数", "只能填写 0-30 的整数")
    _list_validation(ws, "B14", ["自动", "1", "2", "3", "4", "5"], "主表格序号")

    _note_block(
        ws,
        [
            "填写说明",
            "1. 只修改「配置值」列，不要改「配置项」名称。",
            "2. 固定基数单价 / 浮动基数单价按元填写，可填小数。",
            "3. 每页固定行数填 None 时，Word 会按内容自然分页。",
            "4. 单项展示格式中的 {name} 是计费项名称，{amount} 是金额。",
            "5. 输出文件名只能用 {date} 一个占位符。",
            "6. 字体/字号填「自动」时，程序会读取原表格数据行与表头的字体并沿用；",
            "   也可以手动指定中文字体、西文字体和字号（如 宋体 / Times New Roman / 小四）。",
        ],
    )


# ---------------------------------------------------------------------------
# Sheet2 动态计费项
# ---------------------------------------------------------------------------
def _build_items_sheet(ws: Worksheet) -> None:
    ws.sheet_properties.tabColor = "2E75B6"
    _title(ws, "动态计费项（每行一个计费项；计费基数只能填：按次 / 按数量 / 固定金额）", 4)

    headers = ["计费项名称", "单价", "计费基数", "说明"]
    _headers(ws, 2, headers)

    rows = [
        ("值班补贴", 22, "按次", "每行填 1/是/√ 或非零数字，计 1 次"),
        ("绩效奖励", 50, "按数量", "按填写数量 × 单价计费"),
    ]
    for index, (name, price, basis, note) in enumerate(rows, start=3):
        ws.cell(index, 1, name)
        ws.cell(index, 2, price)
        ws.cell(index, 3, basis)
        ws.cell(index, 4, note)
        for col in range(1, 5):
            cell = ws.cell(index, col)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = LEFT if col in (1, 4) else CENTER
            if index % 2 == 0:
                cell.fill = ALT_FILL
        ws.cell(index, 2).number_format = "0.00"
        ws.row_dimensions[index].height = 24

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 46
    ws.freeze_panes = "A3"

    _list_validation(
        ws,
        "C3:C200",
        ["按次", "按数量", "固定金额"],
        "计费基数",
    )
    _note_block(
        ws,
        [
            "填写说明",
            "1. 每行一个计费项，名称不能与标准字段重名。",
            "2. 按次：有值算 1 次；按数量：单价×数量；固定金额：每人固定金额。",
            "3. 源数据中每个计费项一列，列名与计费项名称一致；不一致时可在「字段映射」中配置。",
            "4. 不需要的计费项请删除整行。",
            "5. 示例两行可直接修改，也可删除后新增。",
        ],
        column=6,
    )


# ---------------------------------------------------------------------------
# Sheet3 字段映射
# ---------------------------------------------------------------------------
def _build_mapping_sheet(ws: Worksheet) -> None:
    ws.sheet_properties.tabColor = "548235"
    _title(ws, "字段映射（只填写「原始列名」与「系统标准列名」不一致的列）", 4)

    headers = ["用户原始列名", "系统标准列名", "说明", "示例"]
    _headers(ws, 2, headers)

    rows = [
        (FIELD_NAME, FIELD_NAME, "姓名", "员工姓名 → 姓名"),
        (FIELD_DEPARTMENT, FIELD_DEPARTMENT, "部门", ""),
        (FIELD_EMPLOYEE_ID, FIELD_EMPLOYEE_ID, "工号", "编号 → 工号"),
        (FIELD_FIXED_BASE, FIELD_FIXED_BASE, "固定基数（如工时/天数）", "固定工时 → 固定基数"),
        (FIELD_FLOATING_BASE, FIELD_FLOATING_BASE, "浮动基数（如实际工时）", "实际工时 → 浮动基数"),
        ("值班补贴", "值班补贴", "动态计费项的数据列", ""),
        ("绩效奖励", "绩效奖励", "动态计费项的数据列", ""),
        (FIELD_REMARK, FIELD_REMARK, "备注", ""),
        (FIELD_SEQ, FIELD_SEQ, "序号可留空，程序自动补", ""),
    ]
    for index, (source, target, note, example) in enumerate(rows, start=3):
        ws.cell(index, 1, source)
        ws.cell(index, 2, target)
        ws.cell(index, 3, note)
        ws.cell(index, 4, example)
        for col in range(1, 5):
            cell = ws.cell(index, col)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = LEFT if col != 2 else CENTER
            if index % 2 == 0:
                cell.fill = ALT_FILL
        ws.row_dimensions[index].height = 24

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 46
    ws.column_dimensions["D"].width = 30
    ws.freeze_panes = "A3"

    _note_block(
        ws,
        [
            "填写说明",
            "1. 左列写源数据里的列名，右列写系统标准列名。",
            "2. 系统标准列名必须是标准字段或动态计费项名称。",
            "3. 不需要映射的列可以删除；程序会自动兼容同名列。",
            "4. 改完列名后，请同步检查「源数据.xlsx」的表头。",
        ],
        column=6,
    )


# ---------------------------------------------------------------------------
# 源数据模板
# ---------------------------------------------------------------------------
def _build_data_sheet(ws: Worksheet) -> None:
    ws.title = "源数据"
    ws.sheet_properties.tabColor = "1F4E79"
    _title(ws, "源数据（请保留第 2 行表头；下面两行示例可删除）", 9)

    headers = [
        FIELD_SEQ,
        FIELD_NAME,
        FIELD_DEPARTMENT,
        FIELD_EMPLOYEE_ID,
        FIELD_FIXED_BASE,
        FIELD_FLOATING_BASE,
        "值班补贴",
        "绩效奖励",
        FIELD_REMARK,
    ]
    _headers(ws, 2, headers)

    rows = [
        (1, "张三", "生产部", "A001", 22, 176, 2, 3, "示例数据，可删除"),
        (2, "李四", "质检部", "A002", 20, 160, 1, 0, "示例数据，可删除"),
    ]
    for index, row in enumerate(rows, start=3):
        for col, value in enumerate(row, start=1):
            cell = ws.cell(index, col, value)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = LEFT if col in (2, 3, 9) else CENTER
            if index % 2 == 0:
                cell.fill = ALT_FILL
        ws.cell(index, 1).number_format = "0"
        for col in (5, 6, 7, 8):
            ws.cell(index, col).number_format = "0.##"
        ws.row_dimensions[index].height = 24

    widths = [8, 12, 14, 14, 12, 12, 12, 12, 28]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A3"


def _build_help_sheet(ws: Worksheet) -> None:
    ws.sheet_properties.tabColor = "548235"
    _title(ws, "填写说明", 2)
    lines = [
        ("一、最简流程", "把原始 Word 表格命名为「原版表格.docx」；按本模板填写并保存为「源数据.xlsx」；双击「一键运行.bat」。"),
        ("二、表头", "第 2 行是表头，请保留。列名可以是中文，也可以是英文，只要在「规则配置.xlsx → 字段映射」里能对应上。"),
        ("三、固定/浮动基数", "固定基数与浮动基数是计算基本工资的基础。列名默认为「固定基数」「浮动基数」，可在基础配置里改。"),
        ("四、动态计费项", "「值班补贴」「绩效奖励」只是示例。你可以在「规则配置.xlsx → 动态计费项」里增删，并在本表增加对应列。"),
        ("五、按次与按数量", "按次：填 1、是、√ 或任意非零数字都算 1 次；按数量：填几就按几份计费。"),
        ("六、固定金额", "固定金额型计费项每人金额相同，源数据里可以留空。"),
        ("七、空值", "空单元格按 0 处理；金额支持 1,234.50、￥1234.50、1234.50 元等写法。"),
        ("八、保存", "填写完成后请保存并关闭 Excel，避免文件被占用导致生成失败。"),
        ("九、输出", "运行成功后会在同目录生成「结算报表_YYYYMMDD.docx」。"),
    ]
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 86
    for index, (title, body) in enumerate(lines, start=2):
        ws.cell(index, 1, title).font = Font(name="微软雅黑", size=11, bold=True, color="1F4E79")
        ws.cell(index, 1).alignment = LEFT
        ws.cell(index, 1).border = BORDER
        ws.cell(index, 1).fill = PatternFill("solid", fgColor="DDEBF7")
        ws.cell(index, 2, body).font = BODY_FONT
        ws.cell(index, 2).alignment = LEFT
        ws.cell(index, 2).border = BORDER
        ws.row_dimensions[index].height = 42
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# 通用样式
# ---------------------------------------------------------------------------
def _title(ws: Worksheet, text: str, span: int) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    cell = ws.cell(1, 1, text)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32


def _headers(ws: Worksheet, row: int, headers: list[str]) -> None:
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row, col, text)
        cell.font = WHITE_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[row].height = 26


def _note_block(ws: Worksheet, lines: list[str], column: int = 6) -> None:
    for offset, text in enumerate(lines):
        cell = ws.cell(2 + offset, column, text)
        cell.font = Font(name="宋体", size=9, bold=(offset == 0), color="7F6000")
        cell.fill = NOTE_FILL
        cell.alignment = LEFT
        cell.border = BORDER
        cell.comment = None
    ws.column_dimensions[get_column_letter(column)].width = 52


def _list_validation(
    ws: Worksheet,
    target: str,
    options: list[str],
    title: str,
) -> None:
    formula = '"' + ",".join(options) + '"'
    validation = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
        showDropDown=False,
    )
    validation.errorTitle = f"{title}无效"
    validation.error = "请从下拉列表中选择：" + " / ".join(options)
    validation.promptTitle = title
    validation.prompt = "请选择：" + " / ".join(options)
    validation.showErrorMessage = True
    validation.showInputMessage = True
    ws.add_data_validation(validation)
    validation.add(target)


def _whole_validation(
    ws: Worksheet,
    target: str,
    minimum: int,
    maximum: int,
    title: str,
    message: str,
) -> None:
    validation = DataValidation(
        type="whole",
        operator="between",
        formula1=str(minimum),
        formula2=str(maximum),
        allow_blank=True,
    )
    validation.errorTitle = f"{title}无效"
    validation.error = message
    validation.showErrorMessage = True
    ws.add_data_validation(validation)
    validation.add(target)


__all__ = [
    "CONFIG_FILE",
    "DATA_TEMPLATE_FILE",
    "PROJECT_TITLE",
    "create_config",
    "create_data_template",
]
