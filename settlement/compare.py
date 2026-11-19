# -*- coding: utf-8 -*-
"""与上一期报表做差异对比。

对比基于每次成功运行留下的 ``runs/<run_id>/snapshot.json``，
不依赖重新解析 Word，因此不会受排版变化影响。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import RuleConfig
from .constants import (
    COMPARE_FILE,
    FIELD_EMPLOYEE_ID,
    FIELD_NAME,
    FIELD_SEQ,
    NON_TOTAL_FIELDS,
    RUNS_DIR,
    SNAPSHOT_FILE,
)
from .data import DataTable
from .engine import CalculationResult
from .errors import SettlementError
from .excel import save_workbook_atomic
from .utils import (
    fmt_number,
    info,
    ok,
    read_json,
    to_decimal,
    to_text,
    write_json_atomic,
)


@dataclass(frozen=True)
class FieldChange:
    field: str
    before: str
    after: str


@dataclass
class RowDiff:
    key: str
    label: str
    kind: str  # added / removed / changed
    before: dict[str, str] = field(default_factory=dict)
    after: dict[str, str] = field(default_factory=dict)
    changes: list[FieldChange] = field(default_factory=list)


@dataclass(frozen=True)
class TotalDelta:
    field: str
    before: Decimal
    after: Decimal
    delta: Decimal


@dataclass
class ComparisonResult:
    previous_run: str
    current_run: str
    previous_created_at: str
    current_created_at: str
    rows_before: int
    rows_after: int
    added: list[RowDiff] = field(default_factory=list)
    removed: list[RowDiff] = field(default_factory=list)
    changed: list[RowDiff] = field(default_factory=list)
    unchanged_count: int = 0
    totals: list[TotalDelta] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def added_count(self) -> int:
        return len(self.added)

    @property
    def removed_count(self) -> int:
        return len(self.removed)

    @property
    def changed_count(self) -> int:
        return len(self.changed)

    @property
    def has_changes(self) -> bool:
        return bool(self.added or self.removed or self.changed)

    def summary_text(self) -> str:
        return (
            f"新增 {self.added_count} 行，删除 {self.removed_count} 行，"
            f"修改 {self.changed_count} 行，未变 {self.unchanged_count} 行"
        )


# ---------------------------------------------------------------------------
# 快照
# ---------------------------------------------------------------------------
def build_snapshot(
    *,
    run_id: str,
    data: DataTable,
    calculation: CalculationResult,
    config: RuleConfig,
    pages: int | None,
    padded_rows: int,
    output_name: str | None,
) -> dict[str, Any]:
    """把本次计算结果固化为可比较的行级快照。"""
    rows: list[dict[str, Any]] = []
    seen: dict[str, int] = {}
    for index, row in enumerate(calculation.rows):
        key = _row_key(row.record, index)
        count = seen.get(key, 0) + 1
        seen[key] = count
        if count > 1:
            key = f"{key}#{count}"

        label = _row_label(row.record, index)
        values = {str(name): to_text(value) for name, value in row.record.items()}
        numbers = {
            str(name): fmt_number(value, 6)
            for name, value in row.numbers.items()
            if name not in NON_TOTAL_FIELDS
        }
        rows.append(
            {
                "key": key,
                "label": label,
                "values": values,
                "numbers": numbers,
            }
        )

    totals: dict[str, Decimal] = {}
    for row in calculation.rows:
        for name, value in row.numbers.items():
            if name in NON_TOTAL_FIELDS:
                continue
            totals[name] = totals.get(name, Decimal(0)) + value

    base = config.base
    return {
        "schema": 1,
        "project": "Automated-Settlement-Report-Engine",
        "run_id": run_id,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source": {
            "path": str(data.path),
            "sheet": data.sheet_name,
            "rows": len(calculation.rows),
        },
        "config": {
            "fixed_base_price": str(base.fixed_base_price),
            "floating_base_price": str(base.floating_base_price),
            "rows_per_page": base.rows_per_page,
            "report_title": base.report_title,
            "items": [
                {
                    "name": item.name,
                    "price": str(item.price),
                    "basis": item.basis,
                }
                for item in config.items
            ],
        },
        "output": output_name,
        "row_count": len(calculation.rows),
        "padded_rows": padded_rows,
        "page_count": pages,
        "rows": rows,
        "totals": {name: fmt_number(value, 6) for name, value in totals.items()},
    }


def write_snapshot(root: Path, run_id: str, snapshot: dict[str, Any]) -> Path:
    path = Path(root) / RUNS_DIR / run_id / SNAPSHOT_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    write_json_atomic(path, snapshot)
    return path


def load_snapshot(path: Path) -> dict[str, Any]:
    data = read_json(path)
    if not isinstance(data, dict) or not isinstance(data.get("rows"), list):
        raise SettlementError(f"差异快照格式不正确：{path}")
    return data


def latest_snapshot(
    root: Path,
    *,
    exclude_run_id: str | None = None,
    run_id: str | None = None,
) -> tuple[Path, dict[str, Any]] | None:
    """找上一期快照；默认按 run_id 从新到旧找第一条成功运行。"""
    runs_root = Path(root) / RUNS_DIR
    if run_id:
        path = runs_root / run_id / SNAPSHOT_FILE
        if not path.exists():
            return None
        return path, load_snapshot(path)

    if not runs_root.exists():
        return None
    for directory in sorted(
        (item for item in runs_root.iterdir() if item.is_dir()),
        key=lambda item: item.name,
        reverse=True,
    ):
        if exclude_run_id and directory.name == exclude_run_id:
            continue
        snapshot_path = directory / SNAPSHOT_FILE
        if not snapshot_path.exists():
            continue
        run_json = directory / "run.json"
        if run_json.exists():
            try:
                manifest = read_json(run_json)
            except (OSError, ValueError):
                continue
            if isinstance(manifest, dict) and manifest.get("status") not in (None, "success"):
                continue
        try:
            return snapshot_path, load_snapshot(snapshot_path)
        except (OSError, ValueError, SettlementError):
            continue
    return None


# ---------------------------------------------------------------------------
# 比较
# ---------------------------------------------------------------------------
def compare_snapshots(
    previous: dict[str, Any],
    current: dict[str, Any],
) -> ComparisonResult:
    previous_rows = {str(row.get("key")): row for row in previous.get("rows", [])}
    current_rows = {str(row.get("key")): row for row in current.get("rows", [])}

    added: list[RowDiff] = []
    removed: list[RowDiff] = []
    changed: list[RowDiff] = []
    unchanged = 0

    for key in sorted(current_rows.keys() - previous_rows.keys()):
        row = current_rows[key]
        added.append(
            RowDiff(
                key=key,
                label=str(row.get("label") or key),
                kind="added",
                after=_values(row),
            )
        )

    for key in sorted(previous_rows.keys() - current_rows.keys()):
        row = previous_rows[key]
        removed.append(
            RowDiff(
                key=key,
                label=str(row.get("label") or key),
                kind="removed",
                before=_values(row),
            )
        )

    for key in sorted(current_rows.keys() & previous_rows.keys()):
        before_row = previous_rows[key]
        after_row = current_rows[key]
        before = _values(before_row)
        after = _values(after_row)
        changes: list[FieldChange] = []
        for name in sorted(set(before) | set(after)):
            old = before.get(name, "")
            new = after.get(name, "")
            if _same(old, new):
                continue
            changes.append(FieldChange(field=name, before=old, after=new))
        if changes:
            changed.append(
                RowDiff(
                    key=key,
                    label=str(after_row.get("label") or before_row.get("label") or key),
                    kind="changed",
                    before=before,
                    after=after,
                    changes=changes,
                )
            )
        else:
            unchanged += 1

    totals: list[TotalDelta] = []
    previous_totals = previous.get("totals") or {}
    current_totals = current.get("totals") or {}
    for name in sorted(set(previous_totals) | set(current_totals)):
        before = _decimal(previous_totals.get(name))
        after = _decimal(current_totals.get(name))
        if before != after:
            totals.append(TotalDelta(name, before, after, after - before))

    warnings: list[str] = []
    if previous.get("schema") != current.get("schema"):
        warnings.append("上一期快照版本与本期不同，差异结果仅供参考。")
    if previous.get("config") != current.get("config"):
        warnings.append("两期规则配置不同，金额变化可能来自配置调整。")

    return ComparisonResult(
        previous_run=str(previous.get("run_id") or "?"),
        current_run=str(current.get("run_id") or "?"),
        previous_created_at=str(previous.get("created_at") or ""),
        current_created_at=str(current.get("created_at") or ""),
        rows_before=int(previous.get("row_count") or len(previous_rows)),
        rows_after=int(current.get("row_count") or len(current_rows)),
        added=added,
        removed=removed,
        changed=changed,
        unchanged_count=unchanged,
        totals=totals,
        warnings=warnings,
    )


def write_comparison_report(result: ComparisonResult, path: Path) -> Path:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "差异汇总"
    _write_summary(summary, result)

    rows_sheet = workbook.create_sheet("行差异")
    _write_row_sheet(rows_sheet, result)

    totals_sheet = workbook.create_sheet("合计差异")
    _write_total_sheet(totals_sheet, result)

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    save_workbook_atomic(workbook, path)
    return path


def print_comparison(result: ComparisonResult) -> None:
    info(
        f"差异对比：上期 {result.previous_run} → 本期 {result.current_run}，"
        f"{result.summary_text()}"
    )
    for row in result.added[:10]:
        info(f"  新增：{row.label}")
    for row in result.removed[:10]:
        info(f"  删除：{row.label}")
    for row in result.changed[:10]:
        fields = "、".join(change.field for change in row.changes[:6])
        info(f"  修改：{row.label}（{fields}）")
    for total in result.totals[:10]:
        info(
            f"  合计 {total.field}：{fmt_number(total.before, 6)} → "
            f"{fmt_number(total.after, 6)}（{_signed(total.delta)}）"
        )
    if not result.has_changes and not result.totals:
        ok("与上期没有差异。")


# ---------------------------------------------------------------------------
# 内部辅助
# ---------------------------------------------------------------------------
def _row_key(record: dict[str, Any], index: int) -> str:
    for field in (FIELD_EMPLOYEE_ID, FIELD_NAME, FIELD_SEQ):
        value = to_text(record.get(field)).strip()
        if value:
            return f"{field}={value}"
    return f"row={index + 1}"


def _row_label(record: dict[str, Any], index: int) -> str:
    sequence = to_text(record.get(FIELD_SEQ)).strip()
    name = to_text(record.get(FIELD_NAME)).strip()
    if sequence and name:
        return f"{sequence} {name}"
    if name:
        return name
    if sequence:
        return sequence
    return f"第 {index + 1} 行"


def _values(row: dict[str, Any]) -> dict[str, str]:
    values = row.get("values")
    if not isinstance(values, dict):
        return {}
    return {str(name): to_text(value) for name, value in values.items()}


def _decimal(value: Any) -> Decimal:
    if value is None or str(value).strip() == "":
        return Decimal(0)
    try:
        return to_decimal(value)
    except ValueError:
        return Decimal(0)


def _same(before: str, after: str) -> bool:
    if before == after:
        return True
    try:
        return to_decimal(before) == to_decimal(after)
    except ValueError:
        return False


def _signed(value: Decimal) -> str:
    text = fmt_number(value, 6)
    if value > 0 and not text.startswith("+"):
        return f"+{text}"
    return text


def _write_summary(sheet, result: ComparisonResult) -> None:
    title_font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    label_font = Font(name="微软雅黑", bold=True)
    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet.merge_cells("A1:C1")
    sheet["A1"] = "差异报告"
    sheet["A1"].font = title_font
    sheet.row_dimensions[1].height = 26

    rows = [
        ("上期运行", result.previous_run),
        ("本期运行", result.current_run),
        ("上期时间", result.previous_created_at),
        ("本期时间", result.current_created_at),
        ("上期行数", result.rows_before),
        ("本期行数", result.rows_after),
        ("新增", result.added_count),
        ("删除", result.removed_count),
        ("修改", result.changed_count),
        ("未变", result.unchanged_count),
    ]
    for index, (label, value) in enumerate(rows, start=3):
        sheet.cell(index, 1, label).font = label_font
        sheet.cell(index, 2, value)

    row = 3 + len(rows) + 1
    sheet.cell(row, 1, "合计差异").font = label_font
    row += 1
    for column, header in enumerate(("字段", "上期", "本期", "变化"), start=1):
        cell = sheet.cell(row, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    row += 1
    if result.totals:
        for total in result.totals:
            sheet.cell(row, 1, total.field).border = border
            sheet.cell(row, 2, float(total.before)).border = border
            sheet.cell(row, 3, float(total.after)).border = border
            sheet.cell(row, 4, float(total.delta)).border = border
            row += 1
    else:
        sheet.cell(row, 1, "（没有合计差异）")

    if result.warnings:
        row += 2
        sheet.cell(row, 1, "提示").font = label_font
        row += 1
        for message in result.warnings:
            sheet.cell(row, 1, message)
            row += 1

    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 28
    sheet.column_dimensions["D"].width = 16


def _write_row_sheet(sheet, result: ComparisonResult) -> None:
    header_font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["类型", "键", "标识", "字段", "上期", "本期"]
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    sheet.freeze_panes = "A2"

    row = 2
    for diff in result.added:
        row = _write_diff_row(sheet, row, "新增", diff, "（整行）", "—", "新增", border)
    for diff in result.removed:
        row = _write_diff_row(sheet, row, "删除", diff, "（整行）", "删除", "—", border)
    for diff in result.changed:
        for change in diff.changes:
            row = _write_diff_row(
                sheet,
                row,
                "修改",
                diff,
                change.field,
                change.before,
                change.after,
                border,
            )

    if row == 2:
        sheet.cell(2, 1, "两期没有行级差异。")
    sheet.auto_filter.ref = f"A1:F{max(2, row - 1)}"
    for column, width in zip(range(1, 7), (8, 28, 18, 18, 32, 32)):
        sheet.column_dimensions[get_column_letter(column)].width = width


def _write_diff_row(
    sheet,
    row: int,
    kind: str,
    diff: RowDiff,
    field_name: str,
    before: str,
    after: str,
    border: Border,
) -> int:
    values = [kind, diff.key, diff.label, field_name, before, after]
    for column, value in enumerate(values, start=1):
        cell = sheet.cell(row, column, value)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=(column in (3, 5, 6)))
    return row + 1


def _write_total_sheet(sheet, result: ComparisonResult) -> None:
    header_font = Font(name="微软雅黑", size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["字段", "上期", "本期", "变化"]
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    sheet.freeze_panes = "A2"

    row = 2
    for total in result.totals:
        sheet.cell(row, 1, total.field).border = border
        sheet.cell(row, 2, float(total.before)).border = border
        sheet.cell(row, 3, float(total.after)).border = border
        sheet.cell(row, 4, float(total.delta)).border = border
        row += 1
    if row == 2:
        sheet.cell(2, 1, "两期没有合计差异。")

    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 18
    sheet.column_dimensions["D"].width = 18
