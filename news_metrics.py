"""財政部新聞稿月度指標：只擷取『占累計分配預算數』比率；不以全年達成率替代。

使用方法：
    python crawler.py
    python processor.py
    python news_metrics.py

以 processor 的四表驗證結果為基準，搭配 crawler 的每月新聞稿網址，
輸出 data/news_metrics.csv 與 data/news_metrics_validation.json。
1 月原新聞稿若未公布累計分配預算達成率，保留空值；
12 月保存全年初步實際差額，不假造『累計分配預算』的預估。
"""

import argparse
import csv
import json
import re
import sys
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from bs4 import BeautifulSoup
from crawler import make_session
from update_sheets import load_verified_data

COLUMNS = [
    "資料鍵", "年度", "月份", "資料類型", "官方累計分配預算達成率_%",
    "累計實徵_億元", "全年預算_億元", "預估全年實徵_億元",
    "預估超徵短徵_億元", "全年初步實際差額_億元",
    "預算版本", "新聞網址", "資料狀態", "註記",
]
MONEY_QUANT = Decimal("0.01")

# 僅擷取新聞稿第三段「達成率」的『占累計分配預算數』，不得抓『占全年預算』。
RATIO_PATTERN = re.compile(
    r"占\s*累計\s*分配\s*預算\s*數?\s*(?:達|為)?\s*([0-9]{1,3}(?:\.[0-9]+)?)\s*%"
)
SECTION_PATTERN = re.compile(r"三\s*[、.．]\s*達成率\s*[:：]?", re.S)
PERIOD_PATTERN = re.compile(r"累計\s*1\s*(?:至|到|[-~～])\s*(\d{1,2})\s*月")


def normalized_article_text(html):
    """剔除非新聞本文的 script、nav 等區塊；以文字方式處理 HTML 分段。"""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()
    return unicodedata.normalize("NFKC", soup.get_text(" ", strip=True)).replace("\u3000", " ")


def extract_allocated_ratio(html, year, month):
    """返回 Decimal 百分比或 None；對矛盾格式拋出 ValueError。

    不使用 PDF/Excel 猜測，不取「占全年預算」數字。
    """
    if month == 12:
        return None
    text = normalized_article_text(html)
    matches = list(SECTION_PATTERN.finditer(text))
    if not matches:
        # 1 月的官方新聞稿通常只有『本月數』，沒有第三段。
        if month == 1:
            return None
        raise ValueError(f"{year}/{month:02d} 新聞稿找不到第三段『達成率』")
    if len(matches) != 1:
        raise ValueError(f"{year}/{month:02d} 新聞稿『達成率』出現多次，停止推估")
    end = matches[0].end()
    # 段落一般短於 200 字，限制窗口避免誤抓頁面其他區塊的數字。
    section = text[end:end + 300]
    ratio = RATIO_PATTERN.search(section)
    if not ratio:
        if month == 1:
            return None
        raise ValueError(f"{year}/{month:02d} 找不到『占累計分配預算數』的百分比")
    period = PERIOD_PATTERN.search(section[:ratio.start()])
    if period and int(period.group(1)) != month:
        raise ValueError(f"{year}/{month:02d} 新聞稿累計期間不符：{period.group(1)} 月")
    value = Decimal(ratio.group(1))
    if not (Decimal("1") <= value <= Decimal("500")):
        raise ValueError(f"{year}/{month:02d} 達成率超出合理範圍：{value}")
    return value


def money(x):
    return Decimal(str(x)).quantize(MONEY_QUANT, rounding=ROUND_HALF_UP)


def amount_text(x):
    return format(money(x), ".2f")


def make_metric(year, month, cumulative, budget, version, url, ratio):
    """老師指定公式：全年預算 × 官方分配預算達成率；不是統計預測模型。"""
    cumulative, budget = money(cumulative), money(budget)
    if budget <= 0:
        raise ValueError(f"{year}/{month:02d} 全年預算無效：{budget}")
    if not url.startswith("https://www.mof.gov.tw/"):
        raise ValueError(f"{year}/{month:02d} 新聞稿網址來源異常：{url!r}")
    row = dict.fromkeys(COLUMNS, "")
    row.update({
        "資料鍵": f"{year}{month:02d}", "年度": str(year), "月份": str(month),
        "累計實徵_億元": amount_text(cumulative),
        "全年預算_億元": amount_text(budget),
        "預算版本": version,
        "新聞網址": url,
    })
    if month == 12:
        row.update({
            "資料類型": "全年初步實際", "資料狀態": "初步統計",
            "全年初步實際差額_億元": amount_text(cumulative - budget),
            "註記": "12月採全年初步實徵減全年預算；不屬預估，可能尚未含整理期間入帳。",
        })
    elif ratio is None:
        row.update({
            "資料類型": "未估", "資料狀態": "官方未公布比率",
            "註記": "本月新聞稿未提供累計分配預算達成率；不以全年預算達成率替代。",
        })
    else:
        ratio = Decimal(str(ratio))
        expected = money(budget * ratio / Decimal(100))
        row.update({
            "資料類型": "全年推估", "資料狀態": "依新聞稿比率推估",
            "官方累計分配預算達成率_%": str(ratio),
            "預估全年實徵_億元": amount_text(expected),
            "預估超徵短徵_億元": amount_text(expected - budget),
            "註記": "假設當期相對分配預算達成率延續全年；不是財政部官方預測。",
        })
    return row


def read_manifest(manifest_path):
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    months = manifest.get("months", [])
    if not months:
        raise ValueError("manifest 沒有月份索引")
    return months


def checked_national_totals(rows):
    """只取 6820『總計』官方原始數，避免加總四表重複計算。"""
    totals = {}
    for r in rows:
        if r["來源表"] != "6820" or r["稅目路徑"] != "總計":
            continue
        key = (int(r["年度"]), int(r["月份"]))
        if key in totals:
            raise ValueError(f"{key} 全國總計重複")
        totals[key] = r
    return totals


def build_metrics(manifest_entries, verified_rows, session):
    totals = checked_national_totals(verified_rows)
    entries = {}
    for entry in manifest_entries:
        period = (int(entry["year"]), int(entry["month"]))
        if period in entries:
            raise ValueError(f"重複新聞稿年月：{period}")
        entries[period] = entry
    if set(entries) != set(totals):
        raise ValueError(f"新聞稿索引與驗證成功的稅收年月不一致：{set(entries)^set(totals)}")

    result = []
    for year, month in sorted(entries):
        entry = entries[(year, month)]
        if entry.get("status") != "success":
            raise ValueError(f"{year}/{month:02d} 四表下載不完整，停止新聞稿指標同步")
        url = str(entry.get("news_url", ""))
        if not url.startswith("https://www.mof.gov.tw/"):
            raise ValueError(f"{year}/{month:02d} 新聞稿網址不正確：{url!r}")
        response = session.get(url, timeout=45)
        response.raise_for_status()
        ratio = extract_allocated_ratio(response.text, year, month)
        r = totals[(year, month)]
        result.append(make_metric(
            year, month, r["累計實徵_億元"], r["全年預算_億元"],
            r.get("預算版本", ""), url, ratio,
        ))
        print(f"✓ {year}/{month:02d}：{result[-1]['資料類型']}，比率 {ratio if ratio is not None else '未公布'}")
    return result


def validate_metrics(rows):
    if not rows:
        raise ValueError("沒有新聞稿指標")
    counts = Counter((int(r["年度"]), int(r["月份"])) for r in rows)
    if any(n != 1 for n in counts.values()):
        raise ValueError("新聞稿指標有重複年月")
    latest = max(int(r["年度"]) for r in rows)
    latest_months = {int(r["月份"]) for r in rows if int(r["年度"]) == latest}
    expected = set(range(1, max(latest_months) + 1))
    if latest_months != expected:
        raise ValueError(f"最新年度 {latest} 年月份不連續，缺少 {sorted(expected - latest_months)}")
    # 2~11 月官方每月公告達成率，若網站改版不得靜默輸出錯誤預估。
    for r in rows:
        month = int(r["月份"])
        if 2 <= month <= 11 and r["資料類型"] != "全年推估":
            raise ValueError(f"{r['年度']}/{month:02d} 缺少官方達成率，不可發佈")
        if month == 12 and r["資料類型"] != "全年初步實際":
            raise ValueError("12 月須列全年初步實際，不能當預估")
    return True


def run(manifest_path="data/manifest.json", csv_path="data/tax_records.csv",
        validation_path="data/validation.json", output_dir="data", session=None):
    # 沿用已驗證的三個核心數據與期間完整性檢核。
    verified_rows = load_verified_data(csv_path, validation_path)
    entries = read_manifest(manifest_path)
    if session is None:
        session = make_session()
    metrics = build_metrics(entries, verified_rows, session)
    validate_metrics(metrics)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    dest = output_dir / "news_metrics.csv"
    tmp = dest.with_suffix(".csv.tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(metrics)
    tmp.replace(dest)
    report = {"status": "passed", "months": [r["資料鍵"] for r in metrics],
              "record_count": len(metrics), "estimated_month_count": sum(r["資料類型"] == "全年推估" for r in metrics),
              "source": "Ministry of Finance monthly press releases", "method": "annual_budget * allocated_budget_achievement_ratio"}
    (output_dir / "news_metrics_validation.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n新聞稿指標整理完成：{len(metrics)} 月，已確認的推估月份 {report['estimated_month_count']} 月")
    return metrics


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="擷取財政部新聞稿官方累計分配預算達成率")
    p.add_argument("--manifest", default="data/manifest.json")
    p.add_argument("--csv", default="data/tax_records.csv")
    p.add_argument("--validation", default="data/validation.json")
    p.add_argument("--output", default="data")
    args = p.parse_args()
    try:
        run(args.manifest, args.csv, args.validation, args.output)
    except (ValueError, OSError, KeyError, InvalidOperation) as exc:
        print(f"新聞稿指標驗證失敗：{exc}", file=sys.stderr)
        sys.exit(1)
