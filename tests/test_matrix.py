# -*- coding: utf-8 -*-
"""模板契约、数据校验、运行清单与备份的回归测试。

这些测试刻意使用普通目录（build/test-matrix），不使用 tempfile：
Windows 沙箱会拒绝 Python 创建的临时目录。
"""

from __future__ import annotations

import json
import unittest
import uuid
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

from settlement.config_writer import create_config
from settlement.constants import DATA_CHECK_FILE
from settlement.data import DataTable
from settlement.errors import DataError
from settlement.pipeline import RunOptions, run
from settlement.template import build_template, check_template
from settlement.validation import validate_data, write_validation_report

ROOT = Path(__file__).resolve().parents[1]
WORK_ROOT = ROOT / "build" / "test-matrix"

HEADERS = (
    "序号",
    "姓名",
    "部门",
    "工号",
    "固定基数",
    "浮动基数",
    "值班补贴",
    "绩效奖励",
    "明细",
    "应付合计",
    "备注",
)


class MatrixTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.root = WORK_ROOT / self._testMethodName / uuid.uuid4().hex
        self.root.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # 造数工具
    # ------------------------------------------------------------------
    def _make_source(
        self,
        *,
        header_rows: int = 1,
        footer: bool = True,
        columns: int = 11,
        name: str = "原版表格.docx",
    ) -> Path:
        document = Document()
        total_rows = header_rows + 1 + (1 if footer else 0)
        table = document.add_table(rows=total_rows, cols=columns)
        table.style = "Table Grid"

        for column in range(columns):
            label = HEADERS[column] if column < len(HEADERS) else f"列{column + 1}"
            if header_rows > 1:
                table.cell(0, column).text = "表头"
            table.cell(header_rows - 1, column).text = label

        for column in range(columns):
            table.cell(header_rows, column).text = (
                "张三" if column == 1 else ("10" if column == 4 else "—")
            )

        if footer:
            footer_row = table.rows[-1]
            footer_row.cells[1].text = "合计"
            for column in range(columns):
                if column not in (1, 2, 3):
                    footer_row.cells[column].text = "10"

        path = self.root / name
        document.save(path)
        return path

    def _write_data(self, records: list[dict[str, object]], *, name: str = "源数据.xlsx") -> Path:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(HEADERS)
        for record in records:
            sheet.append([record.get(header, "") for header in HEADERS])
        path = self.root / name
        workbook.save(path)
        return path

    def _valid_record(self) -> dict[str, object]:
        return {
            "序号": 1,
            "姓名": "张三",
            "部门": "技术部",
            "工号": "A001",
            "固定基数": 10,
            "浮动基数": 20,
            "值班补贴": "是",
            "绩效奖励": 3,
            "备注": "测试",
        }

    def _make_config(self) -> Path:
        path = self.root / "规则配置.xlsx"
        create_config(path)
        return path

    # ------------------------------------------------------------------
    # 模板契约
    # ------------------------------------------------------------------
    def test_check_template_accepts_multi_header(self) -> None:
        source = self._make_source(header_rows=2)
        report = check_template(source, header_rows=2)
        self.assertTrue(report.ok, report.errors)
        self.assertEqual(report.header_rows, 2)
        self.assertEqual(report.headers[:2], ["序号", "姓名"])
        self.assertIsNotNone(report.footer_row)

        target = self.root / "_内部模板.docx"
        manifest_path = self.root / "_模板结构.json"
        info = build_template(source, target, manifest_path, header_rows=2)
        self.assertEqual(info.manifest["header_rows"], 2)
        self.assertEqual(info.manifest["data_row_index"], 2)

    def test_check_template_rejects_vertical_merge(self) -> None:
        source = self._make_source()
        document = Document(str(source))
        table = document.tables[0]
        table.cell(1, 0).merge(table.cell(2, 0))
        document.save(source)

        report = check_template(source)
        self.assertFalse(report.ok)
        self.assertTrue(any("纵向合并" in error for error in report.errors), report.errors)

    def test_check_template_rejects_floating_table(self) -> None:
        source = self._make_source()
        document = Document(str(source))
        table = document.tables[0]
        from docx.oxml import OxmlElement

        tbl_pr = table._tbl.tblPr
        tbl_pr.append(OxmlElement("w:tblpPr"))
        document.save(source)

        report = check_template(source)
        self.assertFalse(report.ok)
        self.assertTrue(any("浮动表格" in error for error in report.errors), report.errors)

    def test_check_template_rejects_nested_table(self) -> None:
        source = self._make_source()
        document = Document(str(source))
        table = document.tables[0]
        table.cell(1, 0).add_table(rows=1, cols=1)
        document.save(source)

        report = check_template(source)
        self.assertFalse(report.ok)
        self.assertTrue(any("嵌套表格" in error for error in report.errors), report.errors)

    def test_check_template_warns_horizontal_merge(self) -> None:
        source = self._make_source()
        document = Document(str(source))
        table = document.tables[0]
        table.cell(1, 0).merge(table.cell(1, 1))
        document.save(source)

        report = check_template(source)
        self.assertTrue(report.ok, report.errors)
        self.assertTrue(
            any("横向合并" in warning for warning in report.warnings), report.warnings
        )

    def test_check_template_selects_requested_table(self) -> None:
        document = Document()
        small = document.add_table(rows=1, cols=1)
        small.cell(0, 0).text = "小表"

        table = document.add_table(rows=3, cols=len(HEADERS))
        table.style = "Table Grid"
        for column, header in enumerate(HEADERS):
            table.cell(0, column).text = header
            table.cell(1, column).text = "张三" if column == 1 else "10"
            table.cell(2, column).text = "合计" if column == 1 else "10"

        source = self.root / "多表.docx"
        document.save(source)

        report = check_template(source, table_index=2)
        self.assertTrue(report.ok, report.errors)
        self.assertEqual(report.table_index, 2)

    def test_check_template_warns_when_footer_missing(self) -> None:
        source = self._make_source(footer=False)
        report = check_template(source)
        self.assertTrue(report.ok, report.errors)
        self.assertIsNone(report.footer_row)
        self.assertTrue(any("没有找到" in warning for warning in report.warnings))

    # ------------------------------------------------------------------
    # 数据校验
    # ------------------------------------------------------------------
    def test_validation_report_and_error(self) -> None:
        config = self._make_config()
        from settlement.config import load_config

        table = DataTable(
            path=self.root / "源数据.xlsx",
            sheet_name="数据",
            header_row=1,
            headers=HEADERS,
            fields=HEADERS,
            records=[
                {
                    "序号": "1",
                    "姓名": "张三",
                    "固定基数": "不是数字",
                    "浮动基数": "20",
                }
            ],
        )
        result = validate_data(table, load_config(config))
        self.assertFalse(result.ok)
        self.assertTrue(any(issue.code == "BASE_VALUE_INVALID" for issue in result.errors))

        report_path = self.root / DATA_CHECK_FILE
        write_validation_report(result, report_path)
        workbook = load_workbook(report_path)
        self.assertEqual(workbook.sheetnames, ["检查结果", "问题清单"])
        self.assertEqual(workbook["检查结果"]["B7"].value, "未通过")
        workbook.close()

    def test_invalid_data_blocks_run_and_writes_failed_manifest(self) -> None:
        self._make_config()
        self._write_data([{**self._valid_record(), "固定基数": "不是数字"}])
        self._make_source()

        with self.assertRaises(DataError):
            run(RunOptions(root=self.root))

        self.assertTrue((self.root / DATA_CHECK_FILE).exists())
        manifests = list((self.root / "runs").glob("*/run.json"))
        self.assertEqual(len(manifests), 1)
        payload = json.loads(manifests[0].read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], "failed")
        self.assertIn("DataError", payload["error"])

    # ------------------------------------------------------------------
    # 运行清单与备份
    # ------------------------------------------------------------------
    def test_backup_and_run_manifest_on_second_run(self) -> None:
        self._make_config()
        self._write_data([self._valid_record()])
        self._make_source()

        first = run(RunOptions(root=self.root))
        self.assertTrue(first.output and first.output.exists())
        self.assertTrue(first.run_manifest and first.run_manifest.exists())
        self.assertIsNone(first.backup)
        first_payload = json.loads(first.run_manifest.read_text(encoding="utf-8"))
        self.assertEqual(first_payload["status"], "success")
        self.assertEqual(first_payload["data"]["rows"], 1)
        self.assertTrue(first_payload["result"]["output_sha256"])

        second = run(RunOptions(root=self.root))
        self.assertTrue(second.output and second.output.exists())
        self.assertTrue(second.backup and second.backup.exists())
        self.assertEqual(second.backup.name, second.output.name)
        second_payload = json.loads(second.run_manifest.read_text(encoding="utf-8"))
        self.assertEqual(second_payload["status"], "success")
        self.assertIsNotNone(second_payload["backup"])

    def test_no_validate_skips_report(self) -> None:
        self._make_config()
        self._write_data([self._valid_record()])
        self._make_source()

        result = run(RunOptions(root=self.root, validate=False))
        self.assertIsNone(result.validation)
        self.assertFalse((self.root / DATA_CHECK_FILE).exists())
        self.assertTrue(result.output and result.output.exists())


if __name__ == "__main__":
    unittest.main()
