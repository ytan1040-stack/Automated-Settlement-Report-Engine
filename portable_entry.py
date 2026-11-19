#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""便携版入口：双击 exe 时，以 exe 所在目录作为工作目录。

普通 Python 环境仍可使用 ``python portable_entry.py``；
打包后的 ``一键运行.exe`` 不要求用户安装 Python。
"""

from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    """返回放 原版表格.docx / 源数据.xlsx 的目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main(argv: list[str] | None = None) -> int:
    from settlement.cli import main as cli_main

    return cli_main(argv, default_root=app_root())


if __name__ == "__main__":
    raise SystemExit(main())
