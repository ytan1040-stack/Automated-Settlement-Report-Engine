# -*- coding: utf-8 -*-
"""预览与上期差异对比的集成测试。"""

from __future__ import annotations

import contextlib
import io
import json
import unittest
import uuid
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

from settlement.config_writer import create_config
from settlement.constants import (
    COMPARE_FILE,
    CONFIG_FILE,
    DATA_FILE,
    PREVIEW_FILE,
    SOURCE_DOCX,
)
from settlement.pipeline import RunOptions, run

try:
    import docxtpl  # noqa: F401

    HAS_DOCXTPL = True
except ImportError:  # pragma: no cover
    HAS_DOCXTPL = False


WORK_ROOT = Path(__file__).resolve().parents[1] / "build" / "test-work"
WORK_ROOT.mkdir(parents=True, exist_ok=True)

HEADERS = [
    "序号",
    "姓名",
    "部门",
    "工号",
    "固定基数",
    "浮动基数",
    "值班补贴",
    "绩效奖励",
    "备注",
]
ROW = ["1", "张三", "技术部", "A001", "10", "20", "是", "3", "测试"]


class PreviewCompareTestCase(unittest.TestCase):
    """每个测试使用独立普通目录，避开 Windows 沙箱对 tempfile 的限制。"""

    def setUp(self) -> None:
        self.root = WORK_ROOT / self._testMethodName / uuid.uuid4().hex
        self.root.mkdir(parents=True, exist_ok=True)
        self._make_inputs()

    def _make_inputs(self) -> None:
        document = Document()
        table = document.add_table(rows=3, cols=len(HEADERS))
        for index, header in enumerate(HEADERS):
            table.cell(0, index).text = header
        for index, value in enumerate(ROW):
            table.cell(1, index).text = value
        table.cell(2, 0).text = "合计"
        document.save(self.root / SOURCE_DOCX)

        create_config(self.root / CONFIG_FILE)
        self._write_data()

    def _write_data(self, row: list[str] | None = None) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "源数据"
        worksheet["A1"] = "源数据"
        worksheet.append(HEADERS)
        worksheet.append(row or ROW)
        workbook.save(self.root / DATA_FILE)

    def _run(self, **kwargs):
        with contextlib.redirect_stdout(io.StringIO()):
            return run(RunOptions(root=self.root, **kwargs))

    @unittest.skipUnless(HAS_DOCXTPL, "未安装 docxtpl，跳过预览测试")
    def test_preview_writes_excel_without_word(self) -> None:
        result = self._run(preview=True, preview_rows=5)

        self.assertIsNone(result.output)
        self.assertIsNotNone(result.preview)
        preview_path = self.root / PREVIEW_FILE
        self.assertTrue(preview_path.exists())
        self.assertFalse(list(self.root.glob("结算报表_*.docx")))

        workbook = load_workbook(preview_path, data_only=True)
        self.assertEqual(workbook.sheetnames, ["数据预览", "汇总"])
        sheet = workbook["数据预览"]
        self.assertEqual(sheet.cell(4, 1).value, "页码")
        self.assertEqual(sheet.cell(4, 2).value, "序号")
        self.assertEqual(sheet.cell(5, 2).value, "1")
        self.assertEqual(sheet.cell(5, 3).value, "张三")
        self.assertEqual(sheet.cell(sheet.max_row, 1).value, "合计")

        manifest = self._latest_manifest()
        self.assertIsNotNone(manifest["preview"])
        self.assertIsNone(manifest["result"]["output"])
        self.assertIsNone(manifest["snapshot"])

    @unittest.skipUnless(HAS_DOCXTPL, "未安装 docxtpl，跳过差异测试")
    def test_compare_with_previous_snapshot(self) -> None:
        first = self._run()
        self.assertTrue(first.output.exists())
        first_manifest = self._latest_manifest()
        self.assertIsNotNone(first_manifest["snapshot"])

        # 修改一行：绩效奖励 3 -> 5，金额应增加 100。
        changed = list(ROW)
        changed[7] = "5"
        self._write_data(changed)

        second = self._run(compare=True)
        self.assertIsNotNone(second.comparison)
        self.assertEqual(second.comparison.added_count, 0)
        self.assertEqual(second.comparison.removed_count, 0)
        self.assertEqual(second.comparison.changed_count, 1)

        report_path = self.root / COMPARE_FILE
        self.assertTrue(report_path.exists())
        workbook = load_workbook(report_path, data_only=True)
        self.assertEqual(workbook.sheetnames, ["差异汇总", "行差异", "合计差异"])
        self.assertGreater(workbook["行差异"].max_row, 1)

        manifest = self._latest_manifest()
        self.assertEqual(manifest["compare"]["changed"], 1)
        self.assertEqual(manifest["compare"]["previous_run"], first_manifest["run_id"])
        self.assertIsNotNone(manifest["snapshot"])

    def _latest_manifest(self) -> dict:
        runs = sorted((self.root / "runs").glob("*/run.json"))
        self.assertTrue(runs)
        return json.loads(runs[-1].read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
