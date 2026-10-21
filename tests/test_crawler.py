"""crawler.py 的離線測試：用假的財政部網頁驗證搜尋與連結解析邏輯。

注意：這只驗證程式邏輯，不代表財政部網站可從 GitHub 主機連線。
"""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import crawler  # noqa: E402

FORM = (
    '<form id="form1"><input name="title" value="">'
    '<input name="page" value="1"><input name="token" value="abc"></form>'
)


def results_page(items):
    rows = "".join(
        f'<tr><td><a href="/news/{y}{m:02d}-{i}">{y}年{m}月全國賦稅收入初步統計</a></td></tr>'
        for i, (y, m) in enumerate(items)
    )
    return f'{FORM}<table class="table-list"><tbody>{rows}</tbody></table>'


class FakeResponse:
    def __init__(self, text):
        self.text = text

    def raise_for_status(self):
        pass


class FakeSession:
    def __init__(self, pages):
        self.pages = pages
        self.posted = []

    def get(self, url, **kwargs):
        return FakeResponse(FORM)

    def post(self, url, data=None, **kwargs):
        page = int(dict(data)["page"])
        self.posted.append(page)
        return FakeResponse(self.pages.get(page, results_page([])))


class CrawlerTest(unittest.TestCase):

    def setUp(self):
        self.pages = {
            1: results_page([(115, 9), (115, 8), (115, 7), (115, 6), (115, 5)]),
            2: results_page([(115, 4), (115, 3), (115, 2), (115, 1), (114, 12)]),
            3: results_page([(114, 11), (114, 10), (114, 9), (114, 8), (114, 7)]),
            4: results_page([(114, 6), (114, 5), (114, 4), (114, 3), (114, 2)]),
            5: results_page([(114, 1), (113, 12), (113, 11), (113, 10), (113, 9)]),
            6: results_page([(113, 8)]),
        }

    def test_default_fetches_latest_year_and_previous_year(self):
        session = FakeSession(self.pages)
        articles = crawler.find_articles(session)
        years = {y for y, _ in articles}
        self.assertEqual(years, {114, 115})
        self.assertEqual(len([k for k in articles if k[0] == 115]), 9)
        self.assertEqual(len([k for k in articles if k[0] == 114]), 12)
        # 翻到 113 年那一頁就停止，不會一路翻到第 6 頁
        self.assertEqual(session.posted, [1, 2, 3, 4, 5])

    def test_explicit_year(self):
        session = FakeSession(self.pages)
        articles = crawler.find_articles(session, years=[115])
        self.assertEqual({y for y, _ in articles}, {115})
        # 第 2 頁已出現 114 年，代表 115 年都找完了
        self.assertEqual(session.posted, [1, 2])

    def test_latest_correction_is_kept(self):
        self.pages[1] = results_page([(115, 9), (115, 9)])
        articles = crawler.find_articles(FakeSession(self.pages), years=[115])
        self.assertTrue(articles[(115, 9)]["news_url"].endswith("115" "09-0"))

    def test_no_results_on_first_page_raises(self):
        with self.assertRaises(RuntimeError):
            crawler.find_articles(FakeSession({}), years=[115])

    def test_excel_links_trailing_space_cleaned(self):
        class Session:
            def get(self, url, **kwargs):
                links = "".join(
                    f'<a href="/files/{code}_11509.xlsx%20">{code}</a>'
                    for code in crawler.TABLE_CODES
                )
                return FakeResponse(links)

        article = {"year": 115, "month": 9, "news_url": "https://example.test/n"}
        links = crawler.find_excel_links(Session(), article)
        self.assertEqual(set(links), set(crawler.TABLE_CODES))
        self.assertTrue(all(not u.endswith("%20") for u in links.values()))

    def test_missing_attachment_raises(self):
        class Session:
            def get(self, url, **kwargs):
                return FakeResponse('<a href="/files/6820_11509.xlsx">x</a>')

        article = {"year": 115, "month": 9, "news_url": "https://example.test/n"}
        with self.assertRaises(RuntimeError):
            crawler.find_excel_links(Session(), article)


if __name__ == "__main__":
    unittest.main()
