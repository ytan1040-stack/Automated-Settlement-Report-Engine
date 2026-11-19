# -*- coding: utf-8 -*-
"""面向用户的异常类型。

所有可预期错误都继承 SettlementError，入口脚本只打印简洁的中文提示，
不把 Python 堆栈直接甩给最终用户；设置 SETTLEMENT_DEBUG=1 可查看堆栈。
"""

from __future__ import annotations


class SettlementError(Exception):
    """可预期的业务错误基类。"""


class ConfigError(SettlementError):
    """规则配置.xlsx 缺失、格式错误或取值非法。"""


class DataError(SettlementError):
    """源数据.xlsx 缺失、表头错误或单元格无法解析。"""


class TemplateError(SettlementError):
    """原版表格.docx / _内部模板.docx / 模板结构清单错误。"""


class ReportError(SettlementError):
    """渲染、分页或输出阶段的错误。"""


class FileLockedError(SettlementError):
    """目标文件被 Word / WPS / Excel 占用，无法覆盖。"""

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message
