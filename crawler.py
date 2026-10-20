
"""財政部全國賦稅收入：自動下載每月四張 Excel。"""

import argparse
import json
import re
import ssl
import time
import zipfile

from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter


# ======================================
# 基本設定
# ======================================

SEARCH_URL = (
    "https://www.mof.gov.tw/multiplehtml/"
    "384fb3077bb349ea973e7fc6f13b6974"
)

KEYWORD = "全國賦稅收入初步統計"

TABLE_CODES = (
    "6820",  # 全國稅目
    "6900",  # 政府歸屬
    "6920",  # 中央政府稅目
    "6951",  # 地方政府稅目
)

TITLE_PATTERN = re.compile(
    r"^(\d{2,3})年(\d{1,2})月全國賦稅收入初步統計$"
)


# ======================================
# SSL 相容性
# ======================================

class MOFSSLAdapter(HTTPAdapter):

    def __init__(self, *args, **kwargs):

        self.ssl_context = ssl.create_default_context()

        if hasattr(ssl, "VERIFY_X509_STRICT"):
            self.ssl_context.verify_flags &= (
                ~ssl.VERIFY_X509_STRICT
            )

        super().__init__(*args, **kwargs)

    def init_poolmanager(self, *args, **kwargs):

        kwargs["ssl_context"] = self.ssl_context

        return super().init_poolmanager(
            *args, **kwargs
        )

    def proxy_manager_for(
        self, proxy, **kwargs
    ):

        kwargs["ssl_context"] = self.ssl_context

        if proxy.startswith("https://"):
            kwargs["proxy_ssl_context"] = (
                self.ssl_context
            )

        return super().proxy_manager_for(
            proxy, **kwargs
        )


def make_session():

    session = requests.Session()

    session.mount(
        "https://",
        MOFSSLAdapter()
    )

    session.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "zh-TW,zh;q=0.9",
    })

    return session


# ======================================
# 清理下載網址
# ======================================

def clean_url(raw_url):

    url = raw_url.strip()

    return re.sub(
        r"(?i)(?:%20|%09|%0a|%0d|%c2%a0)+$",
        "",
        url
    )


# ======================================
# 財政部搜尋表單
# ======================================

def build_search_form(html, page):

    soup = BeautifulSoup(
        html, "html.parser"
    )

    form = soup.select_one("form#form1")

    if form is None:
        raise RuntimeError(
            "找不到財政部搜尋表單，"
            "可能是網站改版"
        )

    fields = []

    for control in form.select("input[name]"):

        input_type = control.get(
            "type", "text"
        ).lower()

        if input_type in (
            "submit",
            "button",
            "reset",
            "file"
        ):
            continue

        if input_type in (
            "checkbox",
            "radio"
        ):
            if not control.has_attr("checked"):
                continue

        fields.append((
            control["name"],
            control.get("value", "")
        ))

    overrides = {
        "title": KEYWORD,
        "titleOld": KEYWORD,
        "page": str(page),
        "isPage": (
            "true" if page > 1 else "false"
        ),
        "pageSize": "10",
        "orderField": "ONLINETIME",
        "orderType": "desc",
    }

    present = {
        name for name, _ in fields
    }

    return [
        (
            key,
            overrides.get(key, value)
        )
        for key, value in fields
    ] + [
        (key, value)
        for key, value in overrides.items()
        if key not in present
    ]


# ======================================
# 解析搜尋結果
# ======================================

def parse_news_results(html):

    soup = BeautifulSoup(
        html, "html.parser"
    )

    results = []

    for row in soup.select(
        "table.table-list tbody tr"
    ):

        link = row.select_one("a[href]")

        if link is None:
            continue

        title = re.sub(
            r"\s+",
            "",
            link.get_text(" ", strip=True)
        )

        match = TITLE_PATTERN.fullmatch(
            title
        )

        if not match:
            continue

        year = int(match.group(1))
        month = int(match.group(2))

        if not 1 <= month <= 12:
            continue

        news_url = clean_url(
            urljoin(
                SEARCH_URL,
                link["href"]
            )
        )

        results.append({
            "year": year,
            "month": month,
            "news_url": news_url,
        })

    return results


# ======================================
# 自動搜尋年度新聞
# ======================================

def find_year_articles(
    session,
    year=None,
    max_pages=40
):

    response = session.get(
        SEARCH_URL,
        timeout=40
    )

    response.raise_for_status()

    html = response.text

    articles = {}

    selected_year = year

    for page in range(
        1, max_pages + 1
    ):

        response = session.post(
            SEARCH_URL,
            data=build_search_form(
                html, page
            ),
            headers={
                "Referer": SEARCH_URL
            },
            timeout=40
        )

        response.raise_for_status()

        html = response.text

        matches = parse_news_results(
            html
        )

        print(
            f"搜尋第 {page} 頁："
            f"找到 {len(matches)} 篇統計新聞"
        )

        if not matches:

            if page == 1:
                raise RuntimeError(
                    "搜尋沒有結果，"
                    "請檢查網站是否改版"
                )

            break

        # 不指定年份，自動抓最新年度
        if selected_year is None:
            selected_year = (
                matches[0]["year"]
            )

        for article in matches:

            if (
                article["year"]
                == selected_year
            ):

                month = article["month"]

                articles.setdefault(
                    month, article
                )

        # 搜尋已進入更早年度
        if any(
            article["year"] < selected_year
            for article in matches
        ):
            break

        time.sleep(0.4)

    if not articles:
        raise RuntimeError(
            f"找不到民國 {selected_year} 年"
            "統計新聞稿"
        )

    return selected_year, articles


# ======================================
# 找出新聞稿內的四張 Excel
# ======================================

def find_excel_links(
    session, article
):

    response = session.get(
        article["news_url"],
        timeout=40
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.text, "html.parser"
    )

    suffix = (
        f"{article['year']}"
        f"{article['month']:02d}"
    )

    expected = {
        f"{code}_{suffix}.xlsx": code
        for code in TABLE_CODES
    }

    links = {}

    for tag in soup.select("a[href]"):

        url = clean_url(
            urljoin(
                article["news_url"],
                tag["href"]
            )
        )

        filename = Path(
            unquote(
                urlparse(url).path
            )
        ).name

        if filename in expected:

            code = expected[filename]

            links[code] = url

    missing = (
        set(TABLE_CODES) - set(links)
    )

    if missing:
        raise RuntimeError(
            f"{suffix} 找不到附表："
            + ", ".join(sorted(missing))
        )

    return links


# ======================================
# 下載指定月份 Excel
# ======================================

def download_month(
    session,
    article,
    output_dir
):

    year = article["year"]
    month = article["month"]

    ym = f"{year}{month:02d}"

    folder = (
        output_dir / "excel" / ym
    )

    folder.mkdir(
        parents=True,
        exist_ok=True
    )

    files = {}

    links = find_excel_links(
        session, article
    )

    for code, url in sorted(
        links.items()
    ):

        filename = (
            f"{code}_{ym}.xlsx"
        )

        path = folder / filename

        if (
            not path.is_file()
            or not zipfile.is_zipfile(path)
        ):

            response = session.get(
                url,
                headers={
                    "Referer":
                        article["news_url"]
                },
                timeout=60
            )

            response.raise_for_status()

            path.write_bytes(
                response.content
            )

            if not zipfile.is_zipfile(
                path
            ):

                path.unlink(
                    missing_ok=True
                )

                raise RuntimeError(
                    f"{filename} 不是有效 XLSX"
                )

        files[code] = {
            "path": str(path),
            "url": url
        }

        print(
            f"  ✓ {filename}"
        )

    return files


# ======================================
# 主程式
# ======================================

def run(
    year=None,
    output_dir="data"
):

    output_dir = Path(
        output_dir
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    session = make_session()

    selected_year, articles = (
        find_year_articles(
            session,
            year=year
        )
    )

    print(
        f"\n統計年度："
        f"民國 {selected_year} 年"
    )

    print(
        f"找到月份："
        f"{sorted(articles)}"
    )

    results = []

    for month in sorted(
        articles
    ):

        article = articles[month]

        print(
            f"\n下載 "
            f"{selected_year} 年 "
            f"{month} 月"
        )

        entry = dict(
            article
        )

        try:

            entry["files"] = (
                download_month(
                    session,
                    article,
                    output_dir
                )
            )

            entry["status"] = (
                "success"
            )

        except Exception as exc:

            entry["status"] = (
                "failed"
            )

            entry["error"] = str(
                exc
            )

            print(
                f"  ✗ {exc}"
            )

        results.append(
            entry
        )

        time.sleep(0.4)

    # 儲存下載索引
    manifest = {
        "year": selected_year,
        "updated_at": (
            datetime.now(
                timezone(
                    timedelta(hours=8)
                )
            ).isoformat()
        ),
        "months": results
    }

    manifest_path = (
        output_dir / "manifest.json"
    )

    manifest_path.write_text(
        json.dumps(
            manifest,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    success = sum(
        item["status"] == "success"
        for item in results
    )

    print(
        f"\n完成："
        f"{success}/{len(results)} "
        f"個月份"
    )

    print(
        f"索引：{manifest_path}"
    )

    if success != len(results):
        raise RuntimeError(
            "部分月份下載失敗，"
            "詳見 data/manifest.json；"
            "尚不應更新 Google Sheets"
        )

    return manifest


# ======================================
# GitHub Actions 執行入口
# ======================================

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description=(
            "下載財政部每月賦稅收入"
            "四張 Excel"
        )
    )

    parser.add_argument(
        "--year",
        type=int,
        default=None,
        help=(
            "指定民國年度；"
            "省略則自動抓最新年度"
        )
    )

    parser.add_argument(
        "--output",
        default="data",
        help="輸出資料夾"
    )

    args = parser.parse_args()

    run(
        year=args.year,
        output_dir=args.output
    )
