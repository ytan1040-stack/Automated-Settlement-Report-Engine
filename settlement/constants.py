# -*- coding: utf-8 -*-
"""项目常量与业务字段约定。"""

from __future__ import annotations

PROJECT_NAME = "Automated-Settlement-Report-Engine"
PROJECT_TITLE = "结算报表自动化引擎"
VERSION = "1.0.0"

# ---------------------------------------------------------------------------
# 约定文件名
# ---------------------------------------------------------------------------
CONFIG_FILE = "规则配置.xlsx"
DATA_FILE = "源数据.xlsx"
DATA_TEMPLATE_FILE = "源数据模板.xlsx"
SOURCE_DOCX = "原版表格.docx"
INTERNAL_TEMPLATE = "_内部模板.docx"
MANIFEST_FILE = "_模板结构.json"
DEFAULT_OUTPUT_PATTERN = "结算报表_{date}.docx"
DATA_CHECK_FILE = "数据检查报告.xlsx"
PREVIEW_FILE = "数据预览.xlsx"
COMPARE_FILE = "差异报告.xlsx"
SNAPSHOT_FILE = "snapshot.json"
RUNS_DIR = "runs"
BACKUP_DIR = "backup"

# ---------------------------------------------------------------------------
# 规则配置.xlsx 的工作表
# ---------------------------------------------------------------------------
SHEET_BASE = "基础配置"
SHEET_ITEMS = "动态计费项"
SHEET_MAPPING = "字段映射"

# 生成文件时使用的工作表顺序
CONFIG_SHEET_ORDER = (SHEET_BASE, SHEET_ITEMS, SHEET_MAPPING)

# ---------------------------------------------------------------------------
# Sheet1 配置项名称
# ---------------------------------------------------------------------------
KEY_FIXED_PRICE = "固定基数单价"
KEY_FLOATING_PRICE = "浮动基数单价"
KEY_ITEM_FORMAT = "单项展示格式"
KEY_SEPARATOR = "多项分隔符"
KEY_TOTAL_PREFIX = "总计前缀"
KEY_ROWS_PER_PAGE = "每页固定行数"
KEY_FIXED_COLUMN = "固定基数列"
KEY_FLOATING_COLUMN = "浮动基数列"
KEY_REPORT_TITLE = "报表标题"
KEY_DECIMALS = "金额小数位"
KEY_OUTPUT_NAME = "输出文件名"
KEY_TABLE_INDEX = "主表格序号"
KEY_HEADER_ROWS = "表头行数"
KEY_FONT_EAST_ASIA = "中文字体"
KEY_FONT_ASCII = "西文字体"
KEY_FONT_SIZE = "输出字号"

REQUIRED_BASE_KEYS = (
    KEY_FIXED_PRICE,
    KEY_FLOATING_PRICE,
    KEY_ITEM_FORMAT,
    KEY_SEPARATOR,
    KEY_TOTAL_PREFIX,
    KEY_ROWS_PER_PAGE,
)

DEFAULT_ITEM_FORMAT = "{name}+{amount}"
DEFAULT_SEPARATOR = ","
DEFAULT_TOTAL_PREFIX = ", 共+"
DEFAULT_ROWS_PER_PAGE = 18
DEFAULT_DECIMALS = 2
DEFAULT_FIXED_COLUMN = "固定基数"
DEFAULT_FLOATING_COLUMN = "浮动基数"
DEFAULT_REPORT_TITLE = "结算报表"
DEFAULT_HEADER_ROWS = 1
DEFAULT_FONT_AUTO = "自动"

# ---------------------------------------------------------------------------
# 计费基数
# ---------------------------------------------------------------------------
BASIS_PER_TIME = "按次"
BASIS_PER_QTY = "按数量"
BASIS_FIXED = "固定金额"
BASIS_OPTIONS = (BASIS_PER_TIME, BASIS_PER_QTY, BASIS_FIXED)

# ---------------------------------------------------------------------------
# 标准字段
# ---------------------------------------------------------------------------
FIELD_SEQ = "序号"
FIELD_NAME = "姓名"
FIELD_DEPARTMENT = "部门"
FIELD_EMPLOYEE_ID = "工号"
FIELD_FIXED_BASE = "固定基数"
FIELD_FLOATING_BASE = "浮动基数"
FIELD_BASE_AMOUNT = "基本工资"
FIELD_FIXED_AMOUNT = "固定基数工资"
FIELD_FLOATING_AMOUNT = "浮动基数工资"
FIELD_DETAIL = "明细"
FIELD_DETAIL_TOTAL = "明细合计"
FIELD_TOTAL = "应付合计"
FIELD_REMARK = "备注"

# 计费项名称不能占用的字段（否则会覆盖源数据或系统计算结果）
RESERVED_FIELDS = frozenset(
    {
        FIELD_SEQ,
        FIELD_NAME,
        FIELD_DEPARTMENT,
        FIELD_EMPLOYEE_ID,
        FIELD_FIXED_BASE,
        FIELD_FLOATING_BASE,
        FIELD_BASE_AMOUNT,
        FIELD_FIXED_AMOUNT,
        FIELD_FLOATING_AMOUNT,
        FIELD_DETAIL,
        FIELD_DETAIL_TOTAL,
        FIELD_TOTAL,
        "合计",
        "总计",
        "小计",
        FIELD_REMARK,
    }
)

# 常见的同义列名。只做“没有显式映射时”的兜底，显式字段映射永远优先。
DERIVED_ALIASES = {
    FIELD_SEQ: ("序号", "序列号"),
    FIELD_NAME: ("姓名", "员工姓名", "人员姓名", "名字"),
    FIELD_DEPARTMENT: ("部门", "科室", "所属部门"),
    FIELD_EMPLOYEE_ID: ("工号", "员工编号", "人员编号", "职工号"),
    FIELD_FIXED_BASE: ("固定基数", "固定工时", "基本工时"),
    FIELD_FLOATING_BASE: ("浮动基数", "浮动工时"),
    FIELD_BASE_AMOUNT: ("基本工资", "基础工资", "基数工资"),
    FIELD_DETAIL: ("明细", "明细说明", "补贴说明", "费用明细", "计费明细", "结算明细"),
    FIELD_DETAIL_TOTAL: ("明细合计", "补贴合计", "计费合计", "费用合计"),
    FIELD_TOTAL: (
        "应付合计",
        "合计",
        "总计",
        "应发合计",
        "应发工资",
        "实发合计",
        "实发工资",
        "工资总计",
        "应付金额",
        "结算金额",
    ),
    FIELD_REMARK: ("备注", "说明", "备注信息"),
}

# 这些列即使内容全是数字，也不参与表尾自动求和
NON_TOTAL_FIELDS = frozenset(
    {
        FIELD_SEQ,
        FIELD_NAME,
        FIELD_DEPARTMENT,
        FIELD_EMPLOYEE_ID,
        FIELD_REMARK,
        "日期",
        "时间",
        "身份证号",
        "手机号",
        "电话",
    }
)

# 模板结构清单版本
MANIFEST_VERSION = 1
