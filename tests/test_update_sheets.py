"""update_sheets.py 的測試：用假的工作表驗證型態轉換與覆蓋規則（不連網）。"""

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import processor  # noqa: E402
import update_sheets  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "11509"


class FakeSheet:
    title = "稅收原始資料"

    def __init__(self, header=None, row_count=10, col_count=5, records=None):
        self.header = header or []
        self.records = records or []
        self.row_count = row_count
        self.col_count = col_count
        self.written = []
        self.cleared = []

    def row_values(self, n):
        return list(self.header)

    def get_all_values(self, **kwargs):
        self.value_render_option = kwargs.get("value_render_option")
        if not self.header and not self.records:
            return []
        return [list(self.header)] + [list(row) for row in self.records]

    def resize(self, rows, cols):
        self.row_count, self.col_count = rows, cols

    def update(self, range_name, values, value_input_option):
        self.written.append((range_name, values, value_input_option))

    def batch_clear(self, ranges):
        self.cleared.extend(ranges)

    def freeze(self, rows):
        pass


class UpdateSheetsTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        tmp = tempfile.mkdtemp()
        processor.run(output_dir=tmp, from_dir=str(FIXTURES))
        cls.rows = update_sheets.load_verified_data(
            Path(tmp) / "tax_records.csv", Path(tmp) / "validation.json"
        )

    def test_numbers_stay_numbers(self):
        row = update_sheets.as_sheet_row(self.rows[0])
        by_name = dict(zip(processor.COLUMNS, row))
        self.assertIsInstance(by_name["年度"], int)
        self.assertIsInstance(by_name["累計實徵_億元"], float)
        self.assertIsInstance(by_name["計入加總"], bool)
        self.assertEqual(by_name["累計實徵_億元"], 35713.84)

    def test_empty_rate_is_blank_cell(self):
        row = next(
            r for r in self.rows
            if r["稅目"] == "教育捐" and r["來源表"] == "6820"
        )
        by_name = dict(zip(processor.COLUMNS, update_sheets.as_sheet_row(row)))
        self.assertEqual(by_name["累計占分配預算比率_%"], "")

    def test_sync_overwrites_and_clears_tail(self):
        sheet = FakeSheet(header=processor.COLUMNS)
        count = update_sheets.sync(sheet, self.rows, require_current_contiguous=False)
        self.assertEqual(count, len(self.rows))
        self.assertEqual(sheet.written[0][0], "A1")
        self.assertEqual(sheet.written[0][2], "RAW")
        self.assertEqual(sheet.written[0][1][0], processor.COLUMNS)
        self.assertTrue(sheet.cleared)

    def test_legacy_and_empty_sheets_are_accepted(self):
        update_sheets.sync(FakeSheet(header=update_sheets.LEGACY_COLUMNS), self.rows, require_current_contiguous=False)
        update_sheets.sync(FakeSheet(header=[]), self.rows, require_current_contiguous=False)

    def test_foreign_sheet_is_protected(self):
        with self.assertRaises(ValueError):
            update_sheets.sync(FakeSheet(header=["姓名", "電話"]), self.rows, require_current_contiguous=False)


    def test_previous_year_is_preserved(self):
        base = dict(self.rows[0])
        base["年度"] = "113"
        base["月份"] = "12"
        base["資料鍵"] = "11312|6900|總計"
        preserved = update_sheets.as_sheet_row(base)
        sheet = FakeSheet(header=processor.COLUMNS, records=[preserved])
        count = update_sheets.sync(sheet, self.rows, require_current_contiguous=False)
        self.assertEqual(count, len(self.rows) + 1)
        flattened = [row for _, chunk, _ in sheet.written for row in chunk]
        self.assertIn(preserved, flattened)
        self.assertEqual(sheet.value_render_option, "UNFORMATTED_VALUE")

    def test_partial_latest_year_cannot_publish_as_new_data(self):
        sheet = FakeSheet(header=processor.COLUMNS)
        with self.assertRaisesRegex(ValueError, "缺少"):
            update_sheets.sync(sheet, self.rows)
        self.assertFalse(sheet.written)

    def test_old_legacy_data_is_protected(self):
        sheet = FakeSheet(header=update_sheets.LEGACY_COLUMNS, records=[["old"]])
        with self.assertRaisesRegex(ValueError, "舊版"):
            update_sheets.sync(sheet, self.rows, require_current_contiguous=False)
        self.assertFalse(sheet.written)

    def test_missing_month_override_switch(self):
        import os
        old = os.environ.pop("ALLOW_MISSING_MONTHS", None)
        try:
            self.assertTrue(update_sheets.contiguous_required())
            os.environ["ALLOW_MISSING_MONTHS"] = "1"
            self.assertFalse(update_sheets.contiguous_required())
        finally:
            os.environ.pop("ALLOW_MISSING_MONTHS", None)
            if old is not None:
                os.environ["ALLOW_MISSING_MONTHS"] = old

    def test_current_month_replaces_entire_period_without_duplicate(self):
        old = update_sheets.as_sheet_row(self.rows[0])
        sheet = FakeSheet(header=processor.COLUMNS, records=[old])
        count = update_sheets.sync(sheet, self.rows, require_current_contiguous=False)
        self.assertEqual(count, len(self.rows))



if __name__ == "__main__":
    unittest.main()
