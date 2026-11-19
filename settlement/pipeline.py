# -*- coding: utf-8 -*-
"""端到端流程：读配置 → 读数据 → 建模板 → 计算 → 渲染。"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from .compare import (
    ComparisonResult,
    build_snapshot,
    compare_snapshots,
    latest_snapshot,
    print_comparison,
    write_comparison_report,
    write_snapshot,
)
from .config import RuleConfig, load_config
from .constants import (
    COMPARE_FILE,
    CONFIG_FILE,
    DATA_CHECK_FILE,
    DATA_FILE,
    DEFAULT_FONT_AUTO,
    FIELD_BASE_AMOUNT,
    FIELD_DETAIL,
    FIELD_DETAIL_TOTAL,
    FIELD_FIXED_AMOUNT,
    FIELD_FIXED_BASE,
    FIELD_FLOATING_AMOUNT,
    FIELD_FLOATING_BASE,
    FIELD_SEQ,
    FIELD_TOTAL,
    INTERNAL_TEMPLATE,
    MANIFEST_FILE,
    PROJECT_NAME,
    PROJECT_TITLE,
    SOURCE_DOCX,
    VERSION,
)
from .data import DataTable, load_data
from .engine import CalculationResult, RowCalculation, calculate
from .errors import DataError, SettlementError, TemplateError
from .preview import PreviewResult, print_preview, write_preview_report
from .render import RenderResult, render_report
from .runlog import backup_existing_file, new_run_id, write_run_manifest
from .template import TemplateInfo, build_template, load_manifest
from .validation import ValidationResult, validate_data, write_validation_report
from .utils import (
    banner,
    canonical_field,
    fmt_amount,
    info,
    norm_header,
    ok,
    sha256_file,
    step,
    warn,
)


@dataclass
class RunOptions:
    root: Path
    config_path: Path | None = None
    data_path: Path | None = None
    source_docx: Path | None = None
    template_path: Path | None = None
    manifest_path: Path | None = None
    output_path: Path | None = None
    rebuild_template: bool = False
    sheet_name: str | None = None
    dry_run: bool = False
    validate: bool = True
    backup: bool = True
    preview: bool = False
    preview_rows: int = 10
    compare: bool = False
    compare_run: str | None = None


@dataclass
class RunResult:
    config: RuleConfig
    data: DataTable
    calculation: CalculationResult
    template: TemplateInfo
    output: Path | None
    render: RenderResult | None
    rows: int
    padded_rows: int
    pages: int | None
    warnings: list[str] = field(default_factory=list)
    validation: ValidationResult | None = None
    validation_report: Path | None = None
    run_manifest: Path | None = None
    backup: Path | None = None
    preview: PreviewResult | None = None
    comparison: ComparisonResult | None = None
    comparison_report: Path | None = None
    snapshot: Path | None = None


def run(options: RunOptions) -> RunResult:
    """执行一次完整结算，并留下运行清单。

    所有可预期错误以中文异常抛出；失败时也会尽量写一条 failed 运行记录。
    """
    root = Path(options.root).expanduser().resolve()
    if not root.exists():
        raise SettlementError(f"工作目录不存在：{root}")

    started_at = datetime.now()
    run_id = new_run_id(root, started_at)
    try:
        return _run(options, root, started_at, run_id)
    except Exception as exc:  # noqa: BLE001 - 失败也要留痕，然后原样抛出
        try:
            write_run_manifest(
                root,
                run_id,
                {
                    "schema": 1,
                    "project": PROJECT_NAME,
                    "version": VERSION,
                    "status": "failed",
                    "started_at": started_at,
                    "finished_at": datetime.now(),
                    "root": root,
                    "error": f"{type(exc).__name__}: {exc}",
                },
            )
        except Exception:  # noqa: BLE001 - 运行清单失败不能覆盖原始错误
            pass
        raise


def _run(
    options: RunOptions,
    root: Path,
    started_at: datetime,
    run_id: str,
) -> RunResult:
    config_path = _abs(root, options.config_path, CONFIG_FILE)
    data_path = _abs(root, options.data_path, DATA_FILE)
    source_docx = _abs(root, options.source_docx, SOURCE_DOCX)
    template_path = _abs(root, options.template_path, INTERNAL_TEMPLATE)
    manifest_path = _abs(root, options.manifest_path, MANIFEST_FILE)
    output_path = (
        _abs(root, options.output_path, "") if options.output_path else None
    )

    banner(PROJECT_TITLE, f"{PROJECT_NAME} v{VERSION}")
    info(f"工作目录：{root}")

    step(1, 5, "读取规则配置")
    config = load_config(config_path)
    _print_warnings(config.warnings)
    info(
        f"固定基数单价 {fmt_amount(config.base.fixed_base_price)}，"
        f"浮动基数单价 {fmt_amount(config.base.floating_base_price)}，"
        f"动态计费项 {len(config.items)} 个，"
        + (
            f"每页固定 {config.base.rows_per_page} 行"
            if config.base.rows_per_page
            else "自动分页"
        )
    )

    step(2, 5, "读取源数据")
    header_candidates = list(config.mapping.keys()) + [item.name for item in config.items]
    data = load_data(
        data_path,
        config.mapping,
        sheet_name=options.sheet_name,
        header_candidates=header_candidates,
    )
    _print_warnings(data.warnings)
    info(f"已读取 {data.row_count} 行数据：{data.path.name}（工作表：{data.sheet_name}）")

    validation: ValidationResult | None = None
    validation_report: Path | None = None
    if options.validate:
        validation = validate_data(data, config)
        validation_report = root / DATA_CHECK_FILE
        write_validation_report(validation, validation_report)
        _print_validation(validation)
        if validation.errors:
            raise DataError(
                f"数据检查未通过：{len(validation.errors)} 个错误、"
                f"{len(validation.warnings)} 个警告。"
                f"请查看「{validation_report.name}」的问题清单，修正后重试。"
            )
    else:
        info("已跳过数据校验（--no-validate）")

    step(3, 5, "检查模板")
    template = ensure_template(
        source_docx,
        template_path,
        manifest_path,
        table_index=config.base.table_index,
        header_rows=config.base.header_rows,
        force=options.rebuild_template,
    )
    _print_warnings(template.warnings)

    step(4, 5, "计算")
    calculation = calculate(config, data)
    _print_warnings(calculation.warnings)
    available_fields = _available_fields(config, data)
    render_rows, padding, pages = _pad_rows(
        calculation.rows,
        config.base.rows_per_page,
        available_fields,
    )
    if padding:
        info(f"已补 {padding} 个空行，使每页正好 {config.base.rows_per_page} 行")
    ok(f"计算完成：{len(calculation.rows)} 行 × {len(config.items)} 个计费项")

    step(5, 5, "预览数据" if options.preview else "生成报告")
    today = date.today()
    if output_path is None:
        output_path = root / config.base.output_pattern.format(
            date=today.strftime("%Y%m%d")
        )
    output_path = output_path.resolve()

    context, context_warnings = _build_context(
        config,
        data,
        calculation,
        template.manifest,
        render_rows,
        pages,
    )
    _print_warnings(context_warnings)
    info(_font_summary(config.base, template.manifest))

    snapshot_data = build_snapshot(
        run_id=run_id,
        data=data,
        calculation=calculation,
        config=config,
        pages=pages,
        padded_rows=len(render_rows),
        output_name=output_path.name,
    )

    comparison: ComparisonResult | None = None
    comparison_report: Path | None = None
    if options.compare or options.compare_run:
        previous = latest_snapshot(
            root,
            exclude_run_id=run_id,
            run_id=options.compare_run,
        )
        if previous is None:
            warn(
                "没有找到可对比的上一期快照"
                + (f"（运行 {options.compare_run}）" if options.compare_run else "")
                + "；本期不会生成差异报告。"
            )
        else:
            previous_path, previous_data = previous
            comparison = compare_snapshots(previous_data, snapshot_data)
            comparison_report = root / COMPARE_FILE
            write_comparison_report(comparison, comparison_report)
            print_comparison(comparison)
            info(f"差异报告：{comparison_report.name}（上期 {previous_path.parent.name}）")

    preview_result: PreviewResult | None = None
    if options.preview:
        preview_result = write_preview_report(
            root,
            data,
            calculation,
            config,
            pages=pages,
            padding=padding,
            rows_per_page=config.base.rows_per_page,
        )
        print_preview(preview_result, limit=options.preview_rows)
        warn("预览模式：不写出 Word 文件。")

    render_result = None
    backup_path: Path | None = None
    if options.dry_run or options.preview:
        if options.dry_run and not options.preview:
            warn("试运行模式：不写出 Word 文件。")
    else:
        if options.backup and output_path.exists():
            backup_path = backup_existing_file(output_path, root, run_id)
            info(f"已备份旧输出：{backup_path.relative_to(root)}")
        render_result = render_report(
            template_path,
            output_path,
            context,
            manifest=template.manifest,
            rows_per_page=config.base.rows_per_page,
            expected_rows=len(render_rows),
            font_east_asia=config.base.font_east_asia,
            font_ascii=config.base.font_ascii,
            font_size_pt=config.base.font_size_pt,
        )
        _print_warnings(render_result.warnings)
        ok(f"已生成报告：{output_path.name}")

    snapshot_path: Path | None = None
    if render_result is not None and not options.dry_run and not options.preview:
        snapshot_path = write_snapshot(root, run_id, snapshot_data)

    warnings = _dedupe(
        list(config.warnings)
        + list(data.warnings)
        + list(template.warnings)
        + list(calculation.warnings)
        + list(context_warnings)
        + (list(render_result.warnings) if render_result else [])
        + ([issue.message for issue in validation.warnings] if validation else [])
        + (list(comparison.warnings) if comparison else [])
    )

    final_output = None if (options.dry_run or options.preview) else output_path
    run_manifest = write_run_manifest(
        root,
        run_id,
        {
            "schema": 1,
            "project": PROJECT_NAME,
            "version": VERSION,
            "status": "success",
            "started_at": started_at,
            "finished_at": datetime.now(),
            "root": root,
            "inputs": {
                "config": _file_entry(root, config_path),
                "data": _file_entry(root, data_path),
                "source_docx": _file_entry(root, source_docx),
                "template": _file_entry(root, template_path),
                "template_manifest": _file_entry(root, manifest_path),
            },
            "config": {
                "fixed_base_price": str(config.base.fixed_base_price),
                "floating_base_price": str(config.base.floating_base_price),
                "rows_per_page": config.base.rows_per_page,
                "header_rows": config.base.header_rows,
                "table_index": config.base.table_index,
                "output_pattern": config.base.output_pattern,
                "fonts": {
                    "east_asia": config.base.font_east_asia,
                    "ascii": config.base.font_ascii,
                    "size_pt": config.base.font_size_pt,
                },
            },
            "data": {
                "path": _rel(root, data_path),
                "sheet": data.sheet_name,
                "rows": data.row_count,
                "headers": list(data.headers),
            },
            "template": {
                "path": _rel(root, template_path),
                "source": _rel(root, source_docx),
                "table_index": template.manifest.get("table_index"),
                "header_rows": template.manifest.get("header_rows"),
            },
            "validation": (
                {
                    "status": validation.status,
                    "report": _rel(root, validation_report) if validation_report else None,
                    "errors": len(validation.errors),
                    "warnings": len(validation.warnings),
                }
                if validation
                else {"status": "skipped"}
            ),
            "preview": (
                {
                    "report": _rel(root, preview_result.path),
                    "rows": preview_result.rows,
                    "pages": preview_result.pages,
                }
                if preview_result
                else None
            ),
            "compare": (
                {
                    "previous_run": comparison.previous_run,
                    "current_run": comparison.current_run,
                    "report": _rel(root, comparison_report),
                    "added": comparison.added_count,
                    "removed": comparison.removed_count,
                    "changed": comparison.changed_count,
                    "unchanged": comparison.unchanged_count,
                }
                if comparison
                else None
            ),
            "snapshot": _rel(root, snapshot_path) if snapshot_path else None,
            "result": {
                "output": _rel(root, final_output) if final_output else None,
                "output_sha256": _sha256(final_output) if final_output else None,
                "rows": len(calculation.rows),
                "padded_rows": len(render_rows),
                "pages": pages,
            },
            "backup": _rel(root, backup_path) if backup_path else None,
            "warnings": warnings,
        },
    )
    info(f"运行记录：{run_manifest.relative_to(root)}")
    if validation_report:
        ok(f"数据检查报告：{validation_report.name}")
    else:
        info("数据校验已跳过")

    return RunResult(
        config=config,
        data=data,
        calculation=calculation,
        template=template,
        output=final_output,
        render=render_result,
        rows=len(calculation.rows),
        padded_rows=len(render_rows),
        pages=pages,
        warnings=warnings,
        validation=validation,
        validation_report=validation_report,
        run_manifest=run_manifest,
        backup=backup_path,
        preview=preview_result,
        comparison=comparison,
        comparison_report=comparison_report,
        snapshot=snapshot_path,
    )


# ---------------------------------------------------------------------------
# 模板缓存
# ---------------------------------------------------------------------------
def ensure_template(
    source: Path,
    target: Path,
    manifest_path: Path,
    *,
    table_index: int | None,
    header_rows: int,
    force: bool = False,
) -> TemplateInfo:
    """源表变化或结构参数变化时重建模板；否则复用已有内部模板。"""
    source = Path(source)
    target = Path(target)
    manifest_path = Path(manifest_path)
    warnings: list[str] = []

    if not source.exists():
        if target.exists() and manifest_path.exists():
            warnings.append(
                f"未找到「{source.name}」，继续使用已有的「{target.name}」。"
                "如需重新生成模板，请把原版表格放回程序目录。"
            )
            manifest = load_manifest(manifest_path)
            return TemplateInfo(
                path=target,
                manifest_path=manifest_path,
                manifest=manifest,
                warnings=warnings,
            )
        if target.exists():
            raise TemplateError(
                f"未找到「{source.name}」，且缺少模板结构文件「{manifest_path.name}」，"
                "无法继续使用已有模板。请放回原版表格后重试。"
            )
        raise TemplateError(
            f"缺少「{source.name}」和「{target.name}」。"
            "请把原始 Word 表格命名为「原版表格.docx」放到程序目录，或先运行 make_templates.py。"
        )

    if force:
        info("按要求强制重建模板")
        return build_template(
            source,
            target,
            manifest_path,
            table_index=table_index,
            header_rows=header_rows,
        )

    if not target.exists() or not manifest_path.exists():
        reason = "缺少内部模板" if not target.exists() else "缺少模板结构文件"
        info(f"{reason}，开始生成模板")
        return build_template(
            source,
            target,
            manifest_path,
            table_index=table_index,
            header_rows=header_rows,
        )

    try:
        manifest = load_manifest(manifest_path)
    except TemplateError as exc:
        info(f"模板结构文件不可用（{exc}），重新生成模板")
        return build_template(
            source,
            target,
            manifest_path,
            table_index=table_index,
            header_rows=header_rows,
        )

    rebuild_reason = ""
    if manifest.get("header_rows") != header_rows:
        rebuild_reason = "表头行数已修改"
    elif table_index is not None and manifest.get("table_index") != table_index:
        rebuild_reason = "主表格序号已修改"
    else:
        try:
            current_hash = sha256_file(source)
        except OSError as exc:
            raise TemplateError(f"无法读取「{source.name}」：{exc}") from exc
        if manifest.get("source_sha256") != current_hash:
            rebuild_reason = "原版表格内容已变化"

    if rebuild_reason:
        info(f"{rebuild_reason}，重新生成模板")
        result = build_template(
            source,
            target,
            manifest_path,
            table_index=table_index,
            header_rows=header_rows,
        )
        result.warnings.append(f"{rebuild_reason}，已自动重新生成模板。")
        return result

    info("内部模板已是最新，直接复用")
    return TemplateInfo(
        path=target,
        manifest_path=manifest_path,
        manifest=manifest,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# 渲染上下文
# ---------------------------------------------------------------------------
def _available_fields(config: RuleConfig, data: DataTable) -> set[str]:
    fields = set(data.fields)
    fields.update(
        {
            FIELD_SEQ,
            FIELD_FIXED_BASE,
            FIELD_FLOATING_BASE,
            FIELD_FIXED_AMOUNT,
            FIELD_FLOATING_AMOUNT,
            FIELD_BASE_AMOUNT,
            FIELD_DETAIL,
            FIELD_DETAIL_TOTAL,
            FIELD_TOTAL,
        }
    )
    for item in config.items:
        fields.add(item.name)
        fields.add(f"{item.name}数量")
    return fields


def _build_context(
    config: RuleConfig,
    data: DataTable,
    calculation: CalculationResult,
    manifest: dict[str, Any],
    render_rows: list[RowCalculation],
    pages: int | None,
) -> tuple[dict[str, Any], list[str]]:
    base = config.base
    warnings: list[str] = []
    available = _available_fields_for_context(config, data, calculation)
    data_headers = {
        norm_header(header): field for header, field in zip(data.headers, data.fields)
    }

    headers = list(manifest.get("headers", []))
    slots = list(manifest.get("slots", []))
    header_field: dict[str, str | None] = {}
    for header in headers:
        header_field[header] = _resolve_header(header, data_headers, available)

    unknown = [header for header in headers if header_field.get(header) is None]
    for header in _dedupe(unknown):
        warnings.append(
            f"模板列「{header}」在源数据或计算结果中没有对应字段，将留空。"
            "如需要显示，请在「规则配置.xlsx → 字段映射」中补充映射。"
        )

    row_contexts: list[dict[str, Any]] = []
    for row in render_rows:
        context = dict(row.record)
        for slot, header in zip(slots, headers):
            field = header_field.get(header)
            context[slot] = row.record.get(field, "") if field else ""
        row_contexts.append(context)

    today = date.today()
    context: dict[str, Any] = {
        "rows": row_contexts,
        "report_title": base.report_title,
        "date": today.strftime("%Y-%m-%d"),
        "date_compact": today.strftime("%Y%m%d"),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "row_count": len(calculation.rows),
        "page_count": pages if pages is not None else "自动",
        "rows_per_page": base.rows_per_page if base.rows_per_page else "自动",
        "project": PROJECT_NAME,
        "version": VERSION,
    }

    for entry in manifest.get("total_slots", []):
        slot = entry.get("slot")
        header = entry.get("header", "")
        if not isinstance(slot, int):
            continue
        placeholder = f"total_col{slot}"
        field = header_field.get(header)
        if field is None:
            context[placeholder] = ""
            warnings.append(
                f"表尾合计列「{header}」没有对应数据字段，合计将留空。"
            )
            continue
        total = sum(
            (row.numbers.get(field, Decimal(0)) for row in calculation.rows),
            Decimal(0),
        )
        context[placeholder] = fmt_amount(total, base.decimals)

    return context, _dedupe(warnings)


def _available_fields_for_context(
    config: RuleConfig,
    data: DataTable,
    calculation: CalculationResult,
) -> set[str]:
    fields = _available_fields(config, data)
    if calculation.rows:
        fields.update(calculation.rows[0].record.keys())
    return fields


def _resolve_header(
    header: str,
    data_headers: dict[str, str],
    available: set[str],
) -> str | None:
    key = norm_header(header)
    if key in data_headers:
        return data_headers[key]
    canonical = norm_header(canonical_field(header))
    if canonical in available:
        return canonical
    if key in available:
        return key
    return None


# ---------------------------------------------------------------------------
# 分页补行
# ---------------------------------------------------------------------------
def _pad_rows(
    rows: list[RowCalculation],
    rows_per_page: int | None,
    available_fields: set[str],
) -> tuple[list[RowCalculation], int, int | None]:
    if not rows_per_page:
        return list(rows), 0, None
    remainder = len(rows) % rows_per_page
    padding = (rows_per_page - remainder) % rows_per_page
    blank_record = {field: "" for field in available_fields}
    padded = list(rows) + [
        RowCalculation(record=dict(blank_record), numbers={}) for _ in range(padding)
    ]
    pages = math.ceil(len(padded) / rows_per_page)
    return padded, padding, pages


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------
def _print_validation(validation: ValidationResult) -> None:
    info(
        f"数据检查：{validation.status}"
        f"（错误 {len(validation.errors)}，警告 {len(validation.warnings)}）"
    )
    for issue in validation.issues[:20]:
        location = f"第 {issue.row} 行" if issue.row else "全局"
        warn(f"{location}：{issue.message}")
    if len(validation.issues) > 20:
        warn(f"还有 {len(validation.issues) - 20} 条问题，详见「{DATA_CHECK_FILE}」的问题清单。")


def _file_entry(root: Path, path: Path) -> dict[str, Any]:
    return {
        "path": _rel(root, path),
        "sha256": _sha256(path),
    }


def _rel(root: Path, path: Path | None) -> str | None:
    if path is None:
        return None
    try:
        return str(Path(path).resolve().relative_to(root))
    except (ValueError, OSError):
        return str(path)


def _sha256(path: Path | None) -> str | None:
    if path is None or not Path(path).exists():
        return None
    try:
        return sha256_file(Path(path))
    except OSError:
        return None


def _abs(root: Path, value: Path | str | None, default: str) -> Path:
    if value is None or str(value).strip() == "":
        return (root / default).resolve()
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = root / path
    return path.resolve()


def _print_warnings(warnings: list[str]) -> None:
    for message in _dedupe(warnings):
        warn(message)


def _font_summary(base: Any, manifest: dict[str, Any]) -> str:
    fonts = manifest.get("fonts") or {}
    data_font = fonts.get("data") or {}
    detected = data_font

    auto_east = base.font_east_asia == DEFAULT_FONT_AUTO
    auto_ascii = base.font_ascii == DEFAULT_FONT_AUTO
    auto_size = base.font_size_pt is None

    east = detected.get("east_asia") if auto_east else base.font_east_asia
    ascii_name = detected.get("ascii") if auto_ascii else base.font_ascii
    size = (
        data_font.get("size_pt")
        if auto_size
        else base.font_size_pt
    )

    if auto_east and auto_ascii and auto_size:
        mode = "自动识别"
    elif not auto_east and not auto_ascii and not auto_size:
        mode = "用户指定"
    else:
        mode = "用户指定 + 自动识别"

    east_label = east or "Word 默认"
    ascii_label = ascii_name or "Word 默认"
    size_label = f"{float(size):g} 磅" if size else "跟随模板"
    return f"字体：{mode}（中文 {east_label} / 西文 {ascii_label} / {size_label}）"


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result
