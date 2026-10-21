"""財政部新聞稿之預算達成率離線測試，不依賴即時網站。"""
import sys
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import processor
import news_metrics

SEPT_TEXT = """<html><body><h1>115年9月全國賦稅收入初步統計</h1>
<p>一、本月數：9月實徵淨額4,965億元。</p>
<p>二、累計數：累計1至9月實徵淨額3兆5,714億元。</p>
<p>三、達成率：累計1至9月實徵淨額占累計分配預算數115.0%，占全年預算78.8%。</p>
</body></html>"""
YEAR_END = """<p>一、本月數：12月實徵淨額1,702億元。</p>
<p>二、累計數：114年實徵淨額3兆7,515億元。</p>
<p>三、達成率：114年實徵淨額占全年預算數98.7%，亦即短少505億元。</p>"""
JAN = """<p>115年1月全國賦稅收入初步統計</p>
<p>本月數：1月實徵淨額2,644億元，較上年同期增長。</p>"""


class NewsMetricsTest(unittest.TestCase):
    def test_september_picks_allocated_ratio_not_annual(self):
        self.assertEqual(news_metrics.extract_allocated_ratio(SEPT_TEXT, 115, 9), Decimal('115.0'))

    def test_fullwidth_percent_and_space(self):
        html = '<p>三、達成率：累計1至9月實徵淨額 占累計分配預算數 １１５．０％，占全年預算７８．８％。</p>'
        self.assertEqual(news_metrics.extract_allocated_ratio(html, 115, 9), Decimal('115.0'))

    def test_news_outside_section_not_used(self):
        html = '<p>占累計分配預算數120.0%</p><p>三、達成率：累計1至9月實徵淨額占全年預算78.8%。</p>'
        with self.assertRaisesRegex(ValueError, '找不到'):
            news_metrics.extract_allocated_ratio(html, 115, 9)

    def test_january_no_ratio(self):
        self.assertIsNone(news_metrics.extract_allocated_ratio(JAN, 115, 1))
        r = news_metrics.make_metric(115, 1, '2643.70', '40000', '原預算數',
            'https://www.mof.gov.tw/news?a=1', None)
        self.assertEqual(r['資料類型'], '未估')
        self.assertEqual(r['預估超徵短徵_億元'], '')

    def test_year_end_is_real_preliminary_not_forecast(self):
        self.assertIsNone(news_metrics.extract_allocated_ratio(YEAR_END, 114, 12))
        r = news_metrics.make_metric(114, 12, '37514.80', '38019.39', '原預算數',
            'https://www.mof.gov.tw/news?a=2', None)
        self.assertEqual(r['資料類型'], '全年初步實際')
        self.assertEqual(r['全年初步實際差額_億元'], '-504.59')
        self.assertEqual(r['官方累計分配預算達成率_%'], '')

    def test_september_forecast_teacher_formula(self):
        r = news_metrics.make_metric(115, 9, '35713.84', '45298.76', '依法定數',
            'https://www.mof.gov.tw/news?a=3', Decimal('115.0'))
        self.assertEqual(r['預估全年實徵_億元'], '52093.57')
        self.assertEqual(r['預估超徵短徵_億元'], '6794.81')
        self.assertEqual(r['全年初步實際差額_億元'], '')

    def test_period_mismatch_rejected(self):
        with self.assertRaisesRegex(ValueError, '累計期間不符'):
            news_metrics.extract_allocated_ratio(SEPT_TEXT, 115, 8)

    def test_missing_news_sept_rejected(self):
        with self.assertRaisesRegex(ValueError, '找不到'):
            news_metrics.extract_allocated_ratio(JAN, 115, 9)

    def test_bad_ratio_range_rejected(self):
        html = SEPT_TEXT.replace('115.0%', '900.0%')
        with self.assertRaisesRegex(ValueError, '超出合理範圍'):
            news_metrics.extract_allocated_ratio(html, 115, 9)

    def test_duplicate_month_rejected(self):
        r = news_metrics.make_metric(115, 9, '35713.84', '45298.76', '依法定數',
            'https://www.mof.gov.tw/news?a=3', Decimal('115.0'))
        with self.assertRaisesRegex(ValueError, '重複年月'):
            news_metrics.validate_metrics([r, dict(r)])

    def test_annual_achievement_is_not_allocated_ratio(self):
        html = '<p>三、達成率：累計1至8月實徵淨額占全年預算115.0%。</p>'
        with self.assertRaisesRegex(ValueError, '找不到'):
            news_metrics.extract_allocated_ratio(html, 115, 8)

    def test_budget_and_amount_from_verified_official_6820(self):
        folder = ROOT / 'tests' / 'fixtures' / '11509'
        entry = {'year':115, 'month':9, 'news_url':'https://www.mof.gov.tw/news?a=3',
                 'status':'success', 'files':{code:{'path':str(folder / f'{code}_11509.xlsx'), 'url':''}
                                             for code in processor.TABLES}}
        records,check,warn = processor.parse_month(entry)
        totals = news_metrics.checked_national_totals(records)
        self.assertEqual(len(totals), 1)
        self.assertEqual(totals[(115,9)]['累計實徵_億元'], '35713.84')
        self.assertEqual(totals[(115,9)]['全年預算_億元'], '45298.76')

    def test_complete_offline_three_news_pages(self):
        urls = [f'https://www.mof.gov.tw/news?case={i}' for i in range(3)]
        html_map = dict(zip(urls, [JAN, YEAR_END, SEPT_TEXT]))

        class Response:
            def __init__(self, html):
                self.text = html
            def raise_for_status(self):
                pass

        class Session:
            def get(self, url, timeout):
                return Response(html_map[url])

        periods = [(114, 1), (114, 12), (115, 9)]
        entries = [{"year": y, "month": m, "news_url": u, "status": "success"}
                   for (y, m), u in zip(periods, urls)]
        budgets = [('2194.13', '38019.39'), ('37514.80', '38019.39'),
                   ('35713.84', '45298.76')]
        rows = [{"年度": str(y), "月份": str(m), "來源表": "6820",
                 "稅目路徑": "總計", "累計實徵_億元": cum,
                 "全年預算_億元": budget, "預算版本": "原預算數"}
                for (y, m), (cum, budget) in zip(periods, budgets)]
        result = news_metrics.build_metrics(entries, rows, Session())
        self.assertEqual([r['資料類型'] for r in result],
                         ['未估', '全年初步實際', '全年推估'])
        self.assertEqual(result[-1]['預估超徵短徵_億元'], '6794.81')

    def test_duplicate_achievements_rejected(self):
        html = '<p>三、達成率：累計1至8月實徵淨額占累計分配預算數120.0%。</p>' * 2
        with self.assertRaisesRegex(ValueError, '出現多次'):
            news_metrics.extract_allocated_ratio(html, 115, 8)


if __name__ == '__main__':
    unittest.main()
