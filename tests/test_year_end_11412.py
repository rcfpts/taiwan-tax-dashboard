"""114/12 官方全年版回歸測試，驗證欄位語義與四張表的總額。"""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import processor

FIXTURE = ROOT / "tests" / "fixtures" / "11412"


def month_entry():
    return {
        "year": 114,
        "month": 12,
        "news_url": "",
        "status": "success",
        "files": {
            code: {"path": str(FIXTURE / f"{code}_11412.xlsx"), "url": ""}
            for code in processor.TABLES
        },
    }


class YearEndTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.records, cls.check, cls.warnings = processor.parse_month(month_entry())

    def test_fiscal_year_totals(self):
        self.assertEqual(len(self.records), 82)
        self.assertEqual(self.check["national_cumulative_yi"], "37514.80")
        self.assertEqual(self.check["central_cumulative_yi"], "27362.30")
        self.assertEqual(self.check["local_cumulative_yi"], "8554.49")
        self.assertEqual(self.check["fund_cumulative_yi"], "1598.01")

    def test_december_annual_actual_is_cumulative(self):
        for code in processor.TABLES:
            total = next(
                r for r in self.records
                if r["來源表"] == code and r["稅目路徑"] == "總計"
            )
            self.assertTrue(total["當月實徵_億元"])
            self.assertTrue(total["累計實徵_億元"])
            self.assertTrue(total["全年預算_億元"])
        national = next(r for r in self.records if r["來源表"] == "6820")
        self.assertEqual(national["當月實徵_億元"], "1701.81")
        self.assertEqual(national["累計實徵_億元"], "37514.80")
        self.assertEqual(national["全年預算_億元"], "38019.39")

    def test_regular_and_year_end_headers_both_accepted(self):
        regular = ROOT / "tests" / "fixtures" / "11509" / "6820_11509.xlsx"
        year_end = FIXTURE / "6820_11412.xlsx"
        regular_cells = processor.read_cells(regular)
        december_cells = processor.read_cells(year_end)
        expected = {"month": 2, "cum": 6, "budget": 11}
        self.assertEqual(processor.locate_columns(regular_cells, "test", "6820"), expected)
        self.assertEqual(processor.locate_columns(december_cells, "test", "6820"), expected)

    def test_note_string_in_numeric_cell_is_skipped(self):
        regular = ROOT / "tests" / "fixtures" / "11509" / "6820_11509.xlsx"
        cells = processor.read_cells(regular)
        # 模擬 114/1 尾端跨欄備註：A 欄不是「說明」，但 D 欄是文字。
        cells = {key: value for key, value in cells.items() if key[0] <= 42}
        cells[(43, 0)] = "1.遺產及贈與稅實物抵繳金額1月份計"
        cells[(43, 2)] = "89"
        cells[(43, 3)] = "百萬元，累計1-1月實物抵繳金額共為"
        with patch.object(processor, "read_cells", return_value=cells):
            rows, _ = processor.parse_table("6820", regular, 115, 9)
        self.assertEqual(len(rows), 36)
        self.assertEqual(rows[-1]["name"], "個人")

    def test_other_non_note_garbage_in_tax_row_is_rejected(self):
        regular = ROOT / "tests" / "fixtures" / "11509" / "6820_11509.xlsx"
        cells = processor.read_cells(regular)
        cells[(8, 2)] = "不合法的稅額"
        with patch.object(processor, "read_cells", return_value=cells):
            with self.assertRaisesRegex(ValueError, "無法解析數值"):
                processor.parse_table("6820", regular, 115, 9)


if __name__ == "__main__":
    unittest.main()
