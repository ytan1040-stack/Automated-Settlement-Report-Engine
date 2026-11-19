# -*- coding: utf-8 -*-
"""通用工具：文本/数字转换、金额格式化、日志、路径与文件原子替换。"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import sys
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

from .constants import (
    DERIVED_ALIASES,
    PROJECT_NAME,
    VERSION,
)
from .errors import FileLockedError

DEBUG = os.environ.get("SETTLEMENT_DEBUG", "").strip() == "1"


# ---------------------------------------------------------------------------
# 控制台日志
# ---------------------------------------------------------------------------
def emit(text: str = "") -> None:
    """打印一行日志；极端编码环境下用替换字符兜底，避免日志本身崩溃。"""
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
        print(text.encode(encoding, "replace").decode(encoding, "replace"), flush=True)


def banner(title: str, subtitle: str = "") -> None:
    line = "=" * 68
    emit(line)
    emit(f"  {title}")
    if subtitle:
        emit(f"  {subtitle}")
    emit(line)


def info(message: str) -> None:
    emit(f"[信息] {message}")


def ok(message: str) -> None:
    emit(f"[完成] {message}")


def warn(message: str) -> None:
    emit(f"[警告] {message}")


def error(message: str) -> None:
    emit(f"[错误] {message}")


def hint(message: str) -> None:
    emit(f"[提示] {message}")


def step(index: int, total: int, message: str) -> None:
    emit(f"[{index}/{total}] {message}")


def project_stamp() -> str:
    return f"{PROJECT_NAME} v{VERSION}"


# ---------------------------------------------------------------------------
# 文本与表头
# ---------------------------------------------------------------------------
_WHITESPACE_RE = re.compile(r"\s+", re.UNICODE)


def to_text(value: Any) -> str:
    """把 Excel / Word 单元格值转换成适合展示的单行文本。"""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, Decimal):
        return _decimal_text(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        return _decimal_text(Decimal(str(value)))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def _decimal_text(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def norm_header(value: Any) -> str:
    """表头/字段名归一化：去掉所有空白，便于宽容匹配。"""
    return _WHITESPACE_RE.sub("", to_text(value))


_ALIAS_INDEX: dict[str, str] | None = None
_UNIT_SUFFIX_RE = re.compile(r"[（(]([^（）()]{1,12})[）)]$")
_UNIT_CHARS = set("元万元人次天小时分钟个份项笔台套张间月年日%")


def _strip_unit_suffix(key: str) -> str:
    """去掉「固定基数（元）」「金额(元/月)」这类单位后缀，便于表头匹配。"""
    match = _UNIT_SUFFIX_RE.search(key)
    if not match:
        return key
    inner = match.group(1)
    if inner and all(char in _UNIT_CHARS or char == "/" for char in inner):
        return key[: match.start()]
    return key


def _alias_index() -> dict[str, str]:
    global _ALIAS_INDEX
    if _ALIAS_INDEX is None:
        index: dict[str, str] = {}
        for canonical, aliases in DERIVED_ALIASES.items():
            index[norm_header(canonical)] = canonical
            for alias in aliases:
                index.setdefault(norm_header(alias), canonical)
        _ALIAS_INDEX = index
    return _ALIAS_INDEX


def canonical_field(value: Any) -> str:
    """把常见同义列名归一到标准字段名；无法识别时原样返回。"""
    text = to_text(value)
    if not text:
        return ""
    key = norm_header(text)
    index = _alias_index()
    canonical = index.get(key)
    if canonical:
        return canonical
    stripped = _strip_unit_suffix(key)
    if stripped != key:
        canonical = index.get(stripped)
        if canonical:
            return canonical
    return text


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


# ---------------------------------------------------------------------------
# 数字与金额
# ---------------------------------------------------------------------------
_CURRENCY_CHARS = ("￥", "¥", ",", "，", " ", "\u3000")
_ZERO_PLACEHOLDERS = {"-", "—", "/", "无", "none", "n/a"}


def to_decimal(value: Any) -> Decimal:
    """把单元格值解析为 Decimal。

    支持：整数/浮点、千分位、货币符号、末尾“元”、会计负数括号。
    空值返回 0；无法识别时抛出 ValueError，由调用方补充行号与列名。
    """
    if value is None:
        return Decimal("0")
    if isinstance(value, bool):
        return Decimal(1) if value else Decimal(0)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            raise ValueError(f"不是有效数字：{value!r}")
        return Decimal(str(value))

    text = str(value).strip()
    if not text:
        return Decimal("0")
    for char in _CURRENCY_CHARS:
        text = text.replace(char, "")
    if text.endswith("元"):
        text = text[:-1].strip()
    if text.lower() in _ZERO_PLACEHOLDERS:
        return Decimal("0")

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    if not text:
        raise ValueError("不是有效数字：空白")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"不是有效数字：{value!r}") from exc
    return -number if negative else number


def quantize(value: Any, decimals: int) -> Decimal:
    exponent = Decimal(1).scaleb(-int(decimals))
    return to_decimal(value).quantize(exponent, rounding=ROUND_HALF_UP)


def fmt_amount(value: Any, decimals: int = 2) -> str:
    """金额格式化：固定小数位，不加千分位，避免出现 -0.00。"""
    number = quantize(value, decimals)
    if number == 0:
        number = abs(number)
    return f"{number:.{int(decimals)}f}"


def fmt_number(value: Any, max_places: int = 4) -> str:
    """数量格式化：去掉多余的 0，最多保留 max_places 位小数。"""
    number = quantize(value, max_places)
    if number == 0:
        return "0"
    text = format(number, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


# ---------------------------------------------------------------------------
# 路径与文件
# ---------------------------------------------------------------------------
def resolve_path(base: Path, value: str | Path | None, default: str) -> Path:
    """把相对路径解析到 base 下；value 为空时使用 default。"""
    if value is None or str(value).strip() == "":
        return (base / default).resolve()
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base / path
    return path.resolve()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def write_json_atomic(path: Path, payload: Any) -> None:
    ensure_dir(path.parent)
    temp = path.with_name(f".{path.stem}.{os.getpid()}.tmp.json")
    try:
        with temp.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        replace_file(temp, path)
    finally:
        safe_unlink(temp)


def safe_unlink(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def replace_file(source: Path, target: Path) -> None:
    """用临时文件原子替换目标文件，并在文件被占用时给出可操作的提示。"""
    try:
        os.replace(source, target)
    except PermissionError as exc:
        safe_unlink(source)
        raise FileLockedError(
            f"无法覆盖「{target.name}」：文件可能正在 Word / WPS / Excel 中打开。"
            "请关闭该文件后重试。"
        ) from exc
    except OSError as exc:
        if getattr(exc, "winerror", None) in (5, 32, 33):
            safe_unlink(source)
            raise FileLockedError(
                f"无法覆盖「{target.name}」：文件可能正在 Word / WPS / Excel 中打开。"
                "请关闭该文件后重试。"
            ) from exc
        safe_unlink(source)
        raise


def dedupe(items: Iterable[str]) -> list[str]:
    """按首次出现顺序去重。"""
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result
