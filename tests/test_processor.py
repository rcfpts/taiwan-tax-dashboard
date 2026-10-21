"""processor.py 的回歸測試。

以民國 115 年 9 月的四張財政部附表當固定資料，保護最容易出錯的幾件事：
重複計算、腳註造成名稱不一致、欄位版面改變、--  空值。

執行：python -m unittest discover -s tests -v
"""

import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import processor  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "11509"


def month_entry():
    return {
        "year": 115,
        "month": 9,
        "news_url": "",
        "status": "success",
        "files": {
            code: {"path": str(FIXTURES / f"{code}_11509.xlsx"), "url": ""}
            for code in processor.TABLES
        },
    }


class ProcessorTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.records, cls.check, cls.warnings = processor.parse_month(month_entry())

    def rows(self, code):
        return [r for r in self.records if r["來源表"] == code]

    def find(self, code, path):
        return next(
            r for r in self.rows(code) if r["稅目路徑"] == path
        )

    # ---------- 加總 ----------

    def test_included_rows_sum_to_official_total(self):
        for code in processor.TABLES:
            rows = self.rows(code)
            total = Decimal(rows[0]["累計實徵_億元"])
            leaf = sum(
                Decimal(r["累計實徵_億元"]) for r in rows
                if r["計入加總"] == "TRUE"
            )
            self.assertLessEqual(abs(leaf - total), Decimal("0.05"), code)

    def test_memo_item_is_not_double_counted(self):
        parent = self.find("6820", "房地合一課徵所得稅")
        child = self.find("6820", "房地合一課徵所得稅/個人")
        self.assertEqual(parent["列型態"], "其中項")
        self.assertEqual(parent["計入加總"], "FALSE")
        self.assertEqual(child["列型態"], "其中項明細")
        self.assertEqual(child["計入加總"], "FALSE")

    def test_removing_memo_rule_makes_validation_fail(self):
        original = processor.MEMO_ITEMS
        processor.MEMO_ITEMS = set()
        try:
            with self.assertRaises(ValueError) as ctx:
                processor.parse_month(month_entry())
            self.assertIn("可加總列合計", str(ctx.exception))
        finally:
            processor.MEMO_ITEMS = original

    def test_government_totals(self):
        self.assertEqual(self.check["national_cumulative_yi"], "35713.84")
        self.assertEqual(self.check["central_cumulative_yi"], "23602.67")
        self.assertEqual(self.check["local_cumulative_yi"], "10897.97")
        self.assertEqual(self.check["fund_cumulative_yi"], "1213.21")

    # ---------- 階層與名稱 ----------

    def test_hierarchy(self):
        self.assertEqual(self.find("6820", "所得稅")["列型態"], "小計")
        self.assertEqual(self.find("6820", "所得稅/綜合所得稅")["層級"], 2)
        self.assertEqual(
            self.find("6900", "撥入中央特種基金/長照基金/遺產稅")["政府別"], "基金"
        )

    def test_footnote_marker_removed(self):
        names = {r["稅目"] for r in self.rows("6820")}
        self.assertIn("金融業營業稅(撥入金融業特別準備金)", names)
        self.assertFalse(any(n.endswith("(2)") for n in names))

    def test_keys_unique(self):
        keys = [r["資料鍵"] for r in self.records]
        self.assertEqual(len(keys), len(set(keys)))

    # ---------- 欄位內容 ----------

    def test_yoy_and_ratios_kept(self):
        total = self.rows("6820")[0]
        self.assertEqual(total["累計較上年同期增減_億元"], "7003.84")
        self.assertEqual(total["累計較上年同期增減率_%"], "24.4")
        self.assertEqual(Decimal(total["累計占分配預算比率_%"]), Decimal("115.0"))
        self.assertEqual(total["全年預算_億元"], "45298.76")

    def test_dash_is_empty_not_error(self):
        self.assertIsNone(processor.parse_amount("      --"))
        self.assertIsNone(processor.parse_amount(""))
        self.assertEqual(processor.parse_amount("-4.9"), Decimal("-4.9"))
        row = self.find("6820", "營業稅/金融業營業稅(撥入金融業特別準備金)")
        self.assertEqual(row["當月較上年同月增減率_%"], "")

    def test_budget_version_and_law_version(self):
        self.assertEqual(self.check["budget_version"], "依法定數")
        self.assertTrue(all(r["財劃法版本"] == "新版(115年起)" for r in self.records))

    # ---------- 輸入 ----------

    def test_wrong_period_is_rejected(self):
        entry = month_entry()
        entry["month"] = 8
        with self.assertRaises(ValueError):
            processor.parse_month(entry)

    def test_duplicate_download_filename_still_found(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            for code in processor.TABLES:
                src = FIXTURES / f"{code}_11509.xlsx"
                (tmp / f"abc123-{code}_11509 (1).xlsx").write_bytes(src.read_bytes())
            entries = processor.entries_from_dir(tmp)
            self.assertEqual(len(entries), 1)
            self.assertEqual(set(entries[0]["files"]), set(processor.TABLES))

    def test_run_writes_csv_and_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            processor.run(output_dir=tmp, from_dir=str(FIXTURES))
            self.assertTrue((Path(tmp) / "tax_records.csv").is_file())
            self.assertTrue((Path(tmp) / "validation.json").is_file())


if __name__ == "__main__":
    unittest.main()
