"""印花税流程纯逻辑单元测试。"""
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from config.config_loader import ConfigError
from workflows.jde_flow.main_task import parse_shared_path
from workflows.jde_flow.stamp_tax import (
    _zero_amount,
    build_header_pairs,
    configured_period,
    journal_region,
    month_bounds,
    normalize_header_name,
)
from workflows.jde_flow.excel_com import label_value_coordinate


class PeriodTests(unittest.TestCase):
    def test_month_bounds_from_text(self):
        self.assertEqual(month_bounds("2026-10"), (date(2026, 10, 1), date(2026, 10, 31)))
        self.assertEqual(month_bounds("202602"), (date(2026, 2, 1), date(2026, 2, 28)))

    def test_blank_period_uses_current_month(self):
        today = date.today()
        first, last = month_bounds(configured_period("  "))
        self.assertEqual((first.year, first.month), (today.year, today.month))
        self.assertEqual(first.day, 1)
        self.assertGreaterEqual(last, today - timedelta(days=31))

    def test_invalid_period_rejected(self):
        with self.assertRaises(ConfigError):
            month_bounds("2026-13")


class HeaderMappingTests(unittest.TestCase):
    def test_ignores_spaces_and_separators(self):
        pairs = build_header_pairs(["借项 金额", "帐号"], ["借项金额", "帐号"])
        self.assertEqual(pairs, {"借项 金额": "借项金额", "帐号": "帐号"})

    def test_unmatched_header_left_out(self):
        pairs = build_header_pairs(["帐号", "额外列"], ["帐号"])
        self.assertEqual(pairs, {"帐号": "帐号"})
        self.assertEqual(normalize_header_name(" 借项_金额（元） "), "借项金额元")


class LabelTests(unittest.TestCase):
    def setUp(self):
        self.book = openpyxl.Workbook()
        self.sheet = self.book.active
        self.sheet.title = "印花税计算底稿"

    def tearDown(self):
        self.book.close()

    def test_merged_label_uses_cell_after_merged_range(self):
        self.sheet["B4"] = "会计期间："
        self.sheet.merge_cells("B4:C4")
        self.sheet["B5"] = "批号："
        self.sheet.merge_cells("B5:C5")
        self.assertEqual(label_value_coordinate(self.sheet, "会计期间："), "D4")
        self.assertEqual(label_value_coordinate(self.sheet, "批号："), "D5")

    def test_plain_label_uses_next_cell(self):
        self.sheet["B6"] = "批号："
        self.assertEqual(label_value_coordinate(self.sheet, "批号："), "C6")

    def test_missing_label_rejected(self):
        with self.assertRaises(ConfigError):
            label_value_coordinate(self.sheet, "批号：")


class JournalRegionTests(unittest.TestCase):
    def setUp(self):
        self.book = openpyxl.Workbook()
        self.sheet = self.book.active
        self.sheet["B3"] = "帐号"
        self.sheet["D3"] = "借项\n金额"
        self.sheet["F3"] = "贷项金额"
        self.sheet["H3"] = "备注"
        self.sheet["B4"] = "250104.912302"
        self.sheet["D4"] = 100
        self.sheet["H4"] = "计提"
        self.sheet["D5"] = 0
        self.sheet["B6"] = "250101.400932"
        self.sheet["F6"] = 100
        self.sheet["B7"] = "250105.912302"
        self.sheet["D7"] = 20

    def tearDown(self):
        self.book.close()

    def test_region_starts_after_header_and_ends_at_last_account(self):
        header_row, columns, region = journal_region(self.sheet)
        self.assertEqual(header_row, 3)
        self.assertEqual(columns["账号"], 2)
        self.assertEqual(columns["借项金额"], 4)
        self.assertEqual(columns["贷项金额"], 6)
        self.assertEqual(columns["备注"], 8)
        self.assertEqual(region, [4, 5, 6, 7])

    def test_duplicate_header_rejected(self):
        self.sheet["B10"] = "账号"
        self.sheet["D10"] = "借项金额"
        self.sheet["F10"] = "贷项金额"
        self.sheet["H10"] = "备注"
        with self.assertRaises(ConfigError):
            journal_region(self.sheet)


class AmountTests(unittest.TestCase):
    def test_zero_amount(self):
        for value in (None, "", "  ", 0, 0.0, "0", "0.00"):
            self.assertTrue(_zero_amount(value), value)
        for value in (1, -0.01, "12.3", "abc"):
            self.assertFalse(_zero_amount(value), value)


class SharedPathTests(unittest.TestCase):
    def test_parse_company_and_code(self):
        path, name, code = parse_shared_path("N:\\共享\\太古汇 2501")
        self.assertEqual((name, code), ("太古汇", "2501"))
        self.assertEqual(path.name, "太古汇 2501")

    def test_bad_path_rejected(self):
        with self.assertRaises(ConfigError):
            parse_shared_path("N:\\共享\\太古汇")


if __name__ == "__main__":
    unittest.main()
