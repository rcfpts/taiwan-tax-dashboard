"""穩定性相關功能的離線測試：自動重試、內容相同略過寫入、過期檢查、狀態檔。"""

import json
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import crawler  # noqa: E402
import health  # noqa: E402
import news_metrics  # noqa: E402
import processor  # noqa: E402
import update_news_metrics  # noqa: E402
import update_sheets  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "11509"


class FakeSheet:
    """最小可用的假工作表，記錄所有寫入動作。"""

    def __init__(self, header, records, col_count):
        self.header = header
        self.records = records
        self.row_count = 5000
        self.col_count = col_count
        self.written = []
        self.cleared = []

    def get_all_values(self, **kwargs):
        return [list(self.header)] + [list(r) for r in self.records]

    def resize(self, rows, cols):
        self.row_count, self.col_count = rows, cols

    def update(self, range_name, values, value_input_option):
        self.written.append((range_name, values, value_input_option))

    def batch_clear(self, ranges):
        self.cleared.extend(ranges)

    def freeze(self, rows):
        pass


class RetryTest(unittest.TestCase):

    def test_session_retries_get_and_post_on_server_errors(self):
        adapter = crawler.make_session().get_adapter("https://www.mof.gov.tw/x")
        self.assertIsInstance(adapter, crawler.MOFSSLAdapter)
        retry = adapter.max_retries
        self.assertEqual(retry.total, 3)
        self.assertIn("GET", retry.allowed_methods)
        self.assertIn("POST", retry.allowed_methods)
        self.assertIn(503, retry.status_forcelist)


class SkipUnchangedWriteTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        tmp = tempfile.mkdtemp()
        processor.run(output_dir=tmp, from_dir=str(FIXTURES))
        cls.rows = update_sheets.load_verified_data(
            Path(tmp) / "tax_records.csv", Path(tmp) / "validation.json"
        )

    def _sheet_with(self, rows):
        ordered = sorted(rows, key=update_sheets._sort_key)
        return FakeSheet(
            processor.COLUMNS,
            [update_sheets.as_sheet_row(r) for r in ordered],
            col_count=len(processor.COLUMNS),
        )

    def test_identical_content_is_not_rewritten(self):
        sheet = self._sheet_with(self.rows)
        count = update_sheets.sync(
            sheet, self.rows, require_current_contiguous=False
        )
        self.assertEqual(count, len(self.rows))
        self.assertEqual(sheet.written, [])
        self.assertEqual(sheet.cleared, [])

    def test_changed_value_is_written(self):
        sheet = self._sheet_with(self.rows)
        changed = [dict(r) for r in self.rows]
        changed[0]["累計實徵_億元"] = "1.23"
        update_sheets.sync(sheet, changed, require_current_contiguous=False)
        self.assertTrue(sheet.written)

    def test_removed_row_is_written(self):
        sheet = self._sheet_with(self.rows)
        update_sheets.sync(
            sheet, self.rows[:-1], require_current_contiguous=False
        )
        # 同一年月整月替換：少一列就必須寫入並清掉尾端
        self.assertTrue(sheet.written)
        self.assertTrue(sheet.cleared)


class SkipUnchangedNewsTest(unittest.TestCase):

    @staticmethod
    def _row(ratio):
        return news_metrics.make_metric(
            115, 2, "1000", "1200", "原預算數",
            "https://www.mof.gov.tw/news?a=1", ratio,
        )

    def _sheet(self, records):
        return FakeSheet(
            news_metrics.COLUMNS, records, col_count=len(news_metrics.COLUMNS)
        )

    def test_identical_news_rows_are_not_rewritten(self):
        incoming = [
            news_metrics.make_metric(
                115, 1, "800", "1200", "原預算數",
                "https://www.mof.gov.tw/news?a=0", None,
            ),
            self._row(Decimal("100.6")),
        ]
        sheet = self._sheet(
            [update_news_metrics.to_sheet_row(r) for r in incoming]
        )
        count = update_news_metrics.sync(sheet, incoming)
        self.assertEqual(count, 2)
        self.assertEqual(sheet.written, [])

    def test_changed_ratio_is_written(self):
        jan = news_metrics.make_metric(
            115, 1, "800", "1200", "原預算數",
            "https://www.mof.gov.tw/news?a=0", None,
        )
        stored = [jan, self._row(Decimal("99.0"))]
        incoming = [jan, self._row(Decimal("100.6"))]
        sheet = self._sheet(
            [update_news_metrics.to_sheet_row(r) for r in stored]
        )
        update_news_metrics.sync(sheet, incoming)
        self.assertTrue(sheet.written)


class FreshnessTest(unittest.TestCase):

    def test_expected_latest_before_and_after_grace_day(self):
        self.assertEqual(health.expected_latest(date(2026, 11, 13)), (115, 9))
        self.assertEqual(health.expected_latest(date(2026, 11, 19)), (115, 9))
        self.assertEqual(health.expected_latest(date(2026, 11, 20)), (115, 10))

    def test_expected_latest_crosses_year_boundary(self):
        self.assertEqual(health.expected_latest(date(2027, 1, 5)), (115, 11))
        self.assertEqual(health.expected_latest(date(2027, 1, 20)), (115, 12))
        self.assertEqual(health.expected_latest(date(2027, 2, 1)), (115, 12))

    @staticmethod
    def _validation(year, month):
        return {"months": [
            {"year": year, "month": m} for m in range(1, month + 1)
        ]}

    def test_fresh_data_passes(self):
        ok, _ = health.check_freshness(
            self._validation(115, 9), date(2026, 11, 13)
        )
        self.assertTrue(ok)

    def test_missing_latest_month_fails_after_grace_day(self):
        # 10 月新聞稿因標題改版沒抓到：11/13 還在寬限期，11/20 起就要失敗
        validation = self._validation(115, 9)
        self.assertTrue(
            health.check_freshness(validation, date(2026, 11, 19))[0]
        )
        ok, message = health.check_freshness(validation, date(2026, 11, 20))
        self.assertFalse(ok)
        self.assertIn("115 年 9 月", message)
        self.assertIn("115 年 10 月", message)

    def test_empty_validation_raises(self):
        with self.assertRaises(ValueError):
            health.latest_period({"months": []})


class StatusFileTest(unittest.TestCase):

    def _write_validation(self, folder, year, month):
        path = Path(folder) / "validation.json"
        path.write_text(
            json.dumps({"months": [{"year": year, "month": month}]}),
            encoding="utf-8",
        )
        return path

    def test_status_is_stable_within_a_month(self):
        validation = {"months": [{"year": 115, "month": 9}]}
        a = health.build_status(validation, date(2026, 11, 3))
        b = health.build_status(validation, date(2026, 11, 28))
        self.assertEqual(a, b)
        self.assertEqual(a, {"latest_period": "11509", "checked_month": "2026-11"})

    def test_status_changes_when_month_changes(self):
        validation = {"months": [{"year": 115, "month": 9}]}
        self.assertNotEqual(
            health.build_status(validation, date(2026, 11, 28)),
            health.build_status(validation, date(2026, 12, 1)),
        )

    def test_cli_writes_status_and_check_exit_codes(self):
        with tempfile.TemporaryDirectory() as tmp:
            vpath = self._write_validation(tmp, 115, 9)
            out = Path(tmp) / "status" / "last_run.json"

            code = health.main(
                ["status", "--validation", str(vpath), "--output", str(out)],
                today=date(2026, 11, 13),
            )
            self.assertEqual(code, 0)
            self.assertEqual(
                json.loads(out.read_text(encoding="utf-8")),
                {"latest_period": "11509", "checked_month": "2026-11"},
            )

            self.assertEqual(
                health.main(["check", "--validation", str(vpath)],
                            today=date(2026, 11, 13)),
                0,
            )
            self.assertEqual(
                health.main(["check", "--validation", str(vpath)],
                            today=date(2026, 11, 25)),
                1,
            )


if __name__ == "__main__":
    unittest.main()
