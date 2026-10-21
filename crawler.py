"""財政部全國賦稅收入：自動下載每月四張 Excel（含前一年，供年增與全年推估使用）。"""

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

MONTH_RETRY_DELAYS = (2, 5, 10)  # 斷線時依序等待秒數，最多嘗試四次

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
            self.ssl_context.verify_flags &= ~ssl.VERIFY_X509_STRICT

        super().__init__(*args, **kwargs)

    def init_poolmanager(self, *args, **kwargs):
        kwargs["ssl_context"] = self.ssl_context
        return super().init_poolmanager(*args, **kwargs)

    def proxy_manager_for(self, proxy, **kwargs):
        kwargs["ssl_context"] = self.ssl_context

        if proxy.startswith("https://"):
            kwargs["proxy_ssl_context"] = self.ssl_context

        return super().proxy_manager_for(proxy, **kwargs)


def make_session():
    session = requests.Session()
    session.mount("https://", MOFSSLAdapter())
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
    return re.sub(r"(?i)(?:%20|%09|%0a|%0d|%c2%a0)+$", "", url)


# ======================================
# 財政部搜尋表單
# ======================================

def build_search_form(html, page):
    soup = BeautifulSoup(html, "html.parser")
    form = soup.select_one("form#form1")

    if form is None:
        raise RuntimeError("找不到財政部搜尋表單，可能是網站改版")

    fields = []

    for control in form.select("input[name]"):
        input_type = control.get("type", "text").lower()

        if input_type in ("submit", "button", "reset", "file"):
            continue

        if input_type in ("checkbox", "radio") and not control.has_attr("checked"):
            continue

        fields.append((control["name"], control.get("value", "")))

    overrides = {
        "title": KEYWORD,
        "titleOld": KEYWORD,
        "page": str(page),
        "isPage": "true" if page > 1 else "false",
        "pageSize": "10",
        "orderField": "ONLINETIME",
        "orderType": "desc",
    }

    present = {name for name, _ in fields}

    return [
        (key, overrides.get(key, value)) for key, value in fields
    ] + [
        (key, value) for key, value in overrides.items() if key not in present
    ]


# ======================================
# 解析搜尋結果
# ======================================

def parse_news_results(html):
    soup = BeautifulSoup(html, "html.parser")
    results = []

    for row in soup.select("table.table-list tbody tr"):
        link = row.select_one("a[href]")

        if link is None:
            continue

        title = re.sub(r"\s+", "", link.get_text(" ", strip=True))
        match = TITLE_PATTERN.fullmatch(title)

        if not match:
            continue

        year = int(match.group(1))
        month = int(match.group(2))

        if not 1 <= month <= 12:
            continue

        results.append({
            "year": year,
            "month": month,
            "news_url": clean_url(urljoin(SEARCH_URL, link["href"])),
        })

    return results


# ======================================
# 搜尋要抓的年度（預設：最新年度＋前一年）
# ======================================

def find_articles(session, years=None, max_pages=60):
    """回傳 {(年, 月): 新聞稿}。

    years 省略時，以搜尋結果中最新的年度為準，連同前一年一起抓。
    搜尋結果依發布時間新到舊排列；同一年月若有更正發布，取最新的一篇。
    """

    response = session.get(SEARCH_URL, timeout=40)
    response.raise_for_status()
    html = response.text

    articles = {}
    wanted = set(years) if years else None

    for page in range(1, max_pages + 1):
        response = session.post(
            SEARCH_URL,
            data=build_search_form(html, page),
            headers={"Referer": SEARCH_URL},
            timeout=40,
        )
        response.raise_for_status()
        html = response.text

        matches = parse_news_results(html)

        print(f"搜尋第 {page} 頁：找到 {len(matches)} 篇統計新聞")

        if not matches:
            if page == 1:
                raise RuntimeError("搜尋沒有結果，請檢查網站是否改版")
            break

        if wanted is None:
            latest = matches[0]["year"]
            wanted = {latest, latest - 1}

        for article in matches:
            if article["year"] in wanted:
                articles.setdefault(
                    (article["year"], article["month"]), article
                )

        # 已經翻到比需要的年度更早的新聞，不必再往後找
        if any(a["year"] < min(wanted) for a in matches):
            break

        time.sleep(0.4)

    if not articles:
        raise RuntimeError(f"找不到民國 {sorted(wanted or [])} 年統計新聞稿")

    return articles


# ======================================
# 找出新聞稿內的四張 Excel
# ======================================

def find_excel_links(session, article):
    response = session.get(article["news_url"], timeout=40)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    suffix = f"{article['year']}{article['month']:02d}"

    expected = {f"{code}_{suffix}.xlsx": code for code in TABLE_CODES}

    links = {}

    for tag in soup.select("a[href]"):
        url = clean_url(urljoin(article["news_url"], tag["href"]))
        filename = Path(unquote(urlparse(url).path)).name

        if filename in expected:
            links[expected[filename]] = url

    missing = set(TABLE_CODES) - set(links)

    if missing:
        raise RuntimeError(f"{suffix} 找不到附表：" + ", ".join(sorted(missing)))

    return links


# ======================================
# 下載指定月份 Excel
# ======================================

def download_month(session, article, output_dir):
    ym = f"{article['year']}{article['month']:02d}"

    folder = output_dir / "excel" / ym
    folder.mkdir(parents=True, exist_ok=True)

    files = {}

    for code, url in sorted(find_excel_links(session, article).items()):
        filename = f"{code}_{ym}.xlsx"
        path = folder / filename

        if not path.is_file() or not zipfile.is_zipfile(path):
            response = session.get(
                url,
                headers={"Referer": article["news_url"]},
                timeout=60,
            )
            response.raise_for_status()

            path.write_bytes(response.content)

            if not zipfile.is_zipfile(path):
                path.unlink(missing_ok=True)
                raise RuntimeError(f"{filename} 不是有效 XLSX")

        files[code] = {"path": str(path), "url": url}

        print(f"  ✓ {filename}")

    return files


# ======================================
# 主程式
# ======================================

def run(years=None, output_dir="data"):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    session = make_session()

    articles = find_articles(session, years=years)

    all_years = sorted({year for year, _ in articles})
    latest_year = max(all_years)

    print(f"\n統計年度：民國 {all_years} 年")

    results = []

    for (year, month) in sorted(articles):
        article = articles[(year, month)]

        print(f"\n下載 {year} 年 {month} 月")

        entry = dict(article)

        # 同一月份遇到網路暫時斷線時重試；已下載成功的 XLSX 會沿用。
        # 重要：不把下載失敗誤寫成「該月份沒有數據」。
        for attempt in range(1, len(MONTH_RETRY_DELAYS) + 2):
            try:
                entry["files"] = download_month(session, article, output_dir)
                entry["status"] = "success"
                break
            except requests.exceptions.RequestException as exc:
                if attempt > len(MONTH_RETRY_DELAYS):
                    entry["status"] = "failed"
                    entry["error"] = str(exc)
                    print(f"  ✗ 重試後仍下載失敗：{exc}")
                    break
                delay = MONTH_RETRY_DELAYS[attempt - 1]
                print(f"  ⚠ 網路暫時中斷（第 {attempt} 次），{delay} 秒後重試：{exc}")
                time.sleep(delay)
            except Exception as exc:
                # 例如新聞稿已改版、附件缺少：不可當成臨時斷線或正常缺值。
                entry["status"] = "failed"
                entry["error"] = str(exc)
                print(f"  ✗ {exc}")
                break

        results.append(entry)

        time.sleep(0.4)

    manifest = {
        "years": all_years,
        "updated_at": datetime.now(timezone(timedelta(hours=8))).isoformat(),
        "months": results,
    }

    manifest_path = output_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    success = sum(item["status"] == "success" for item in results)

    print(f"\n完成：{success}/{len(results)} 個月份")
    print(f"索引：{manifest_path}")

    # 最新年度有月份失敗才視為失敗；前一年缺的月份由 processor 以警告處理
    failed_current = [
        item for item in results
        if item["status"] != "success" and item["year"] == latest_year
    ]

    if failed_current:
        raise RuntimeError(
            "最新年度部分月份下載失敗，詳見 data/manifest.json；"
            "可改用 data/inbox/ 手動放入 Excel"
        )

    return manifest


# ======================================
# GitHub Actions 執行入口
# ======================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="下載財政部每月賦稅收入四張 Excel")

    parser.add_argument(
        "--years",
        type=int,
        nargs="*",
        default=None,
        help="指定民國年度（可多個）；省略則抓最新年度與前一年",
    )

    parser.add_argument("--output", default="data", help="輸出資料夾")

    args = parser.parse_args()

    run(years=args.years or None, output_dir=args.output)
