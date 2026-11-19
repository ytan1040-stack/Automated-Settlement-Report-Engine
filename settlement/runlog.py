# -*- coding: utf-8 -*-
"""运行清单与覆盖前备份。

每次正式运行都会在 ``runs/<时间戳>/run.json`` 留下可追溯记录；
如果输出文件已存在，会先复制到 ``backup/<时间戳>/``。
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from .constants import BACKUP_DIR, RUNS_DIR
from .errors import FileLockedError
from .utils import write_json_atomic


def new_run_id(root: Path, started_at: datetime | None = None) -> str:
    """生成不会冲突的运行编号，例如 20261119-231500。"""
    root = Path(root)
    base = (started_at or datetime.now()).strftime("%Y%m%d-%H%M%S")
    candidate = base
    suffix = 1
    while (root / RUNS_DIR / candidate).exists() or (root / BACKUP_DIR / candidate).exists():
        suffix += 1
        candidate = f"{base}-{suffix}"
    return candidate


def backup_existing_file(path: Path, root: Path, run_id: str) -> Path | None:
    """覆盖前备份输出文件；不存在时返回 None。"""
    path = Path(path)
    if not path.exists():
        return None
    target = Path(root) / BACKUP_DIR / run_id / path.name
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(path, target)
    except PermissionError as exc:
        raise FileLockedError(
            f"无法备份「{path.name}」：文件可能正在 Word / WPS 中打开。"
            "请关闭文件后重试。"
        ) from exc
    except OSError as exc:
        raise FileLockedError(f"无法备份「{path.name}」：{exc}") from exc
    return target


def write_run_manifest(root: Path, run_id: str, payload: dict[str, Any]) -> Path:
    """写出 run.json，返回路径。"""
    path = Path(root) / RUNS_DIR / run_id / "run.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {"run_id": run_id, **payload}
    write_json_atomic(path, _jsonable(data))
    return path


def _jsonable(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, datetime):
        return value.astimezone().isoformat(timespec="seconds")
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if hasattr(value, "as_tuple"):  # Decimal 等
        return str(value)
    return value
