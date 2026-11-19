# -*- coding: utf-8 -*-
"""最小但覆盖关键路径的单元/集成测试。

运行：
    python -m unittest discover -s tests -v

集成测试需要 docxtpl；如果未安装，会跳过。
"""

from __future__ import annotations

import contextlib
import io
import unittest
import uuid
from decimal import Decimal
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import Workbook, load_workbook

from settlement.config import BaseConfig, BillingItem, RuleConfig, load_config
from settlement.config_writer import create_config, create_data_template
from settlement.constants import (
    CONFIG_FILE,
    DATA_FILE,
    INTERNAL_TEMPLATE,
    MANIFEST_FILE,
    SOURCE_DOCX,
)
from settlement.data import DataTable
from settlement.engine import calculate
from settlement.pipeline import RunOptions, run
from settlement.utils import canonical_field, fmt_amount, to_decimal

try:
    import docxtpl  # noqa: F401

    HAS_DOCXTPL = True
except ImportError:  # pragma: no cover - 仅在未安装依赖时触发
    HAS_DOCXTPL = False


WORK_ROOT = Path(__file__).resolve().parents[1] / "build" / "test-work"
WORK_ROOT.mkdir(parents=True, exist_ok=True)


class WorkspaceTestCase(unittest.TestCase):
    """每个测试使用独立的普通目录。

    不使用 tempfile.TemporaryDirectory：在本项目的 Windows 沙箱里，它创建的
    目录会被拒绝写入/清理；普通 Path.mkdir 则正常。build/ 已被 .gitignore。
    """

    def setUp(self) -> None:
        self.root = WORK_ROOT / self._testMethodName / uuid.uuid4().hex
        self.root.mkdir(parents=True, exist_ok=True)


HEADERS = [
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
]
ROW = ["1", "张三", "技术部", "A001", "10", "20", "是", "3", "", "", "测试"]


class MoneyTests(unittest.TestCase):
    def test_parse_money(self) -> None:
        self.assertEqual(to_decimal("￥1,234.50元"), Decimal("1234.50"))
        self.assertEqual(to_decimal("(1,234.50)"), Decimal("-1234.50"))
        self.assertEqual(to_decimal(" - "), Decimal("0"))

    def test_format_money(self) -> None:
        self.assertEqual(fmt_amount(Decimal("802"), 2), "802.00")
        self.assertEqual(fmt_amount(Decimal("-0.004"), 2), "0.00")

    def test_header_aliases_and_units(self) -> None:
        self.assertEqual(canonical_field("员工姓名"), "姓名")
        self.assertEqual(canonical_field("固定基数（元）"), "固定基数")
        self.assertEqual(canonical_field("备注（必填）"), "备注（必填）")


class ConfigTests(WorkspaceTestCase):
    def setUp(self) -> None:
        super().setUp()

    def test_roundtrip(self) -> None:
        config_path = self.root / CONFIG_FILE
        create_config(config_path)
        config = load_config(config_path)

        self.assertEqual(config.base.fixed_base_price, Decimal("15"))
        self.assertEqual(config.base.floating_base_price, Decimal("24"))
        self.assertEqual(config.base.rows_per_page, 18)
        self.assertEqual(config.items[0].name, "值班补贴")
        self.assertEqual(config.mapping["姓名"], "姓名")
        self.assertEqual(config.mapping["值班补贴"], "值班补贴")
        self.assertEqual(config.base.font_east_asia, "自动")
        self.assertEqual(config.base.font_ascii, "自动")
        self.assertIsNone(config.base.font_size_pt)

    def test_data_template_shape(self) -> None:
        template_path = self.root / "源数据模板.xlsx"
        create_data_template(template_path)
        workbook = load_workbook(template_path, data_only=True)
        worksheet = workbook.active
        self.assertEqual(worksheet.max_column, 9)
        self.assertEqual(worksheet.max_row, 4)
        self.assertEqual(worksheet.cell(2, 1).value, "序号")
        self.assertEqual(worksheet.cell(2, 9).value, "备注")


class EngineTests(WorkspaceTestCase):
    def setUp(self) -> None:
        super().setUp()
        config_path = self.root / CONFIG_FILE
        create_config(config_path)
        self.config = load_config(config_path)

    def test_calculation_and_detail(self) -> None:
        table = DataTable(
            path=self.root / "源数据.xlsx",
            sheet_name="源数据",
            header_row=2,
            headers=HEADERS,
            fields=HEADERS,
            records=[dict(zip(HEADERS, ROW))],
            warnings=[],
        )
        result = calculate(self.config, table)
        record = result.rows[0].record

        self.assertEqual(record["基本工资"], "630.00")
        self.assertEqual(record["明细合计"], "172.00")
        self.assertEqual(record["应付合计"], "802.00")
        self.assertEqual(
            record["明细"],
            "值班补贴+22.00,绩效奖励+150.00, 共+172.00",
        )
        self.assertEqual(result.rows[0].numbers["应付合计"], Decimal("802.00"))

    def test_negative_detail_does_not_double_sign(self) -> None:
        config = RuleConfig(
            path=self.root / CONFIG_FILE,
            base=BaseConfig(
                fixed_base_price=Decimal("0"),
                floating_base_price=Decimal("0"),
                item_format="{name}+{amount}",
                separator=",",
                total_prefix=", 共+",
                rows_per_page=None,
                fixed_base_column="固定基数",
                floating_base_column="浮动基数",
                report_title="负数测试",
                decimals=2,
                output_pattern="结算报表_{date}.docx",
                table_index=None,
                header_rows=1,
            ),
            items=(BillingItem("数量", Decimal("1"), "按数量"),),
            mapping={},
        )
        table = DataTable(
            path=self.root / "源数据.xlsx",
            sheet_name="源数据",
            header_row=2,
            headers=["数量"],
            fields=["数量"],
            records=[{"数量": "-24"}],
            warnings=[],
        )

        record = calculate(config, table).rows[0].record
        self.assertEqual(record["明细"], "数量-24.00, 共-24.00")
        self.assertEqual(record["应付合计"], "-24.00")


@unittest.skipUnless(HAS_DOCXTPL, "未安装 docxtpl，跳过端到端测试")
class EndToEndTests(WorkspaceTestCase):
    def setUp(self) -> None:
        super().setUp()
        self._make_source_docx()
        self._make_workbooks()

    def _make_source_docx(self) -> None:
        document = Document()
        table = document.add_table(rows=3, cols=len(HEADERS))
        for index, header in enumerate(HEADERS):
            table.cell(0, index).text = header
        for index, value in enumerate(ROW):
            table.cell(1, index).text = value
        table.cell(2, 0).text = "合计"
        for cell in table.rows[0].cells:
            _set_cell_font(cell, "黑体", "Times New Roman", 14)
        for cell in table.rows[1].cells:
            _set_cell_font(cell, "仿宋_GB2312", "Times New Roman", 12)
        document.save(self.root / SOURCE_DOCX)

    def _make_workbooks(self) -> None:
        create_config(self.root / CONFIG_FILE)

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "源数据"
        worksheet["A1"] = "源数据"
        worksheet.append(HEADERS)
        worksheet.append(ROW)
        workbook.save(self.root / DATA_FILE)

    def test_full_pipeline(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            result = run(RunOptions(root=self.root))

        self.assertIsNotNone(result.output)
        self.assertTrue(result.output.exists())
        self.assertTrue((self.root / INTERNAL_TEMPLATE).exists())
        self.assertTrue((self.root / MANIFEST_FILE).exists())
        self.assertEqual(result.rows, 1)
        self.assertEqual(result.padded_rows, 18)
        self.assertEqual(result.pages, 1)

        document = Document(result.output)
        self.assertEqual(len(document.tables), 1)
        table = document.tables[0]
        self.assertEqual(len(table.rows), 20)  # 表头 + 18 数据行（17 空行）+ 合计

        data_row = [cell.text for cell in table.rows[1].cells]
        self.assertEqual(data_row[1], "张三")
        self.assertEqual(data_row[8], "值班补贴+22.00,绩效奖励+150.00, 共+172.00")
        self.assertEqual(data_row[9], "802.00")

        footer = [cell.text for cell in table.rows[-1].cells]
        self.assertEqual(footer[0], "合计")
        self.assertEqual(footer[4], "10.00")
        self.assertEqual(footer[5], "20.00")
        self.assertEqual(footer[6], "22.00")
        self.assertEqual(footer[7], "150.00")
        self.assertEqual(footer[9], "802.00")
        self.assertNotIn("{{", document.element.xml)
        self.assertNotIn("{%", document.element.xml)


    def test_font_auto_and_override(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            first = run(RunOptions(root=self.root))

        document = Document(first.output)
        table = document.tables[0]
        east, ascii_name, size = _first_run_font(table.rows[0].cells[0])
        self.assertEqual(east, "黑体")
        self.assertEqual(ascii_name, "Times New Roman")
        self.assertEqual(size, 14)
        east, ascii_name, size = _first_run_font(table.rows[1].cells[0])
        self.assertEqual(east, "仿宋_GB2312")
        self.assertEqual(ascii_name, "Times New Roman")
        self.assertEqual(size, 12)

        workbook = load_workbook(self.root / CONFIG_FILE)
        worksheet = workbook["基础配置"]
        for row_index in range(1, worksheet.max_row + 1):
            key = worksheet.cell(row_index, 1).value
            if key in {"中文字体", "西文字体", "输出字号"}:
                worksheet.cell(row_index, 2).value = {
                    "中文字体": "宋体",
                    "西文字体": "Arial",
                    "输出字号": 10.5,
                }[key]
        workbook.save(self.root / CONFIG_FILE)

        output = self.root / "自定义字体.docx"
        with contextlib.redirect_stdout(io.StringIO()):
            run(RunOptions(root=self.root, output_path=output))

        document = Document(output)
        table = document.tables[0]
        for row_index in (0, 1):
            east, ascii_name, size = _first_run_font(table.rows[row_index].cells[0])
            self.assertEqual(east, "宋体")
            self.assertEqual(ascii_name, "Arial")
            self.assertEqual(size, 10.5)


def _set_cell_font(cell, east_asia: str, ascii_name: str, size_pt: float) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = ascii_name
            run.font.size = Pt(size_pt)
            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.get_or_add_rFonts()
            r_fonts.set(qn("w:eastAsia"), east_asia)


def _first_run_font(cell) -> tuple[str | None, str | None, float | None]:
    run = cell.paragraphs[0].runs[0]
    r_pr = run._element.rPr
    r_fonts = r_pr.rFonts if r_pr is not None else None
    east = r_fonts.get(qn("w:eastAsia")) if r_fonts is not None else None
    ascii_name = r_fonts.get(qn("w:ascii")) if r_fonts is not None else None
    size = run.font.size.pt if run.font.size is not None else None
    return east, ascii_name, size


if __name__ == "__main__":
    unittest.main()
