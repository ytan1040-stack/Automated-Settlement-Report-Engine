#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""读取源数据并生成结算报表。

最简用法：
    python make_report.py

程序默认读取同级目录的：
    规则配置.xlsx、源数据.xlsx、_内部模板.docx、_模板结构.json

输出：
    结算报表_YYYYMMDD.docx
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import docx  # noqa: F401
    import docxtpl  # noqa: F401
    import openpyxl  # noqa: F401
except ImportError:
    print("缺少运行依赖（python-docx / docxtpl / openpyxl）。")
    print("请先执行：python -m pip install -r requirements.txt")
    raise SystemExit(1)

from settlement.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main(default_root=ROOT))
