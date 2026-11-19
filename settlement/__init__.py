# -*- coding: utf-8 -*-
"""Automated Settlement Report Engine 核心包。

本包只包含与具体业务无关的结算报表能力：
配置读取、源数据读取、计费、Word 模板生成、渲染与分页。
"""

from .constants import PROJECT_NAME, VERSION

__all__ = ["PROJECT_NAME", "VERSION"]
