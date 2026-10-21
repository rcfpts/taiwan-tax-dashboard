"""獨立新聞稿工作表的離線安全寫入測試。"""
import sys
import unittest
from pathlib import Path
from decimal import Decimal

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import news_metrics
import update_news_metrics


def row(year, month, ratio):
    return news_metrics.make_metric(year, month, '1000', '1200', '原預算數',
        'https://www.mof.gov.tw/news?a=1', ratio)


class FakeSheet:
    def __init__(self, records=None, header=None):
        self.header = header or []
        self.records = records or []
        self.row_count, self.col_count = 50, len(news_metrics.COLUMNS)
        self.written = []
        self.cleared = []
        self.frozen = None

    def get_all_values(self, **kwargs):
        if not self.header:
            return []
        return [self.header] + [list(r) for r in self.records]

    def resize(self, rows, cols):
        self.row_count = rows
        self.col_count = cols

    def update(self, range_name, values, value_input_option):
        self.written.append((range_name, values, value_input_option))

    def batch_clear(self, ranges):
        self.cleared.extend(ranges)

    def freeze(self, rows):
        self.frozen = rows


class UpdateNewsMetricsTest(unittest.TestCase):
    def test_inserts_and_retains_past_year(self):
        old = update_news_metrics.to_sheet_row(row(113,12,None))
        sheet = FakeSheet([old], news_metrics.COLUMNS)
        incoming = [row(115,1,None), row(115,2,Decimal('100.6'))]
        count = update_news_metrics.sync(sheet, incoming)
        self.assertEqual(count, 3)
        values = sheet.written[0][1]
        self.assertEqual(values[0], news_metrics.COLUMNS)
        self.assertIn(old, values)
        self.assertEqual(sheet.frozen, 1)

    def test_no_duplicate_month_on_repeat_sync(self):
        current = row(115,1,None)
        existing = update_news_metrics.to_sheet_row(current)
        sheet = FakeSheet([existing], news_metrics.COLUMNS)
        count = update_news_metrics.sync(sheet,[current])
        self.assertEqual(count,1)

    def test_unsafe_existing_header_refuses_writing(self):
        sheet = FakeSheet([], ['電話','金額'])
        with self.assertRaisesRegex(ValueError, '表頭不同'):
            update_news_metrics.sync(sheet, [row(115,1,None)])
        self.assertFalse(sheet.written)

    def test_number_and_empty_cells(self):
        r = row(115,2,Decimal('100.6'))
        values = dict(zip(news_metrics.COLUMNS, update_news_metrics.to_sheet_row(r)))
        self.assertEqual(values['年度'],115)
        self.assertEqual(values['官方累計分配預算達成率_%'],100.6)
        self.assertEqual(values['全年初步實際差額_億元'],'')

    def test_missing_middle_month_blocks_publish(self):
        sheet = FakeSheet()
        with self.assertRaisesRegex(ValueError, '缺少'):
            update_news_metrics.sync(sheet,[row(115,1,None),row(115,3,Decimal('111.0'))])
        self.assertFalse(sheet.written)

    def test_december_forecast_disallowed(self):
        fake = row(114,12,None)
        fake['資料類型'] = '全年推估'
        full_year = [row(114,1,None)] + [row(114,m,Decimal('105.0')) for m in range(2,12)] + [fake]
        with self.assertRaisesRegex(ValueError, '12 月'):
            news_metrics.validate_metrics(full_year)


if __name__ == '__main__':
    unittest.main()
