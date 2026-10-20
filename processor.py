
"""財政部月度賦稅 Excel：整理及驗證。

執行：
    python crawler.py
    python processor.py

讀取：
    data/manifest.json
    data/excel/<年月>/<代碼>_<年月>.xlsx

輸出：
    data/tax_records.csv
    data/validation.json
"""

import argparse
import csv
import json
import re
import zipfile

from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "x": (
        "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main"
    )
}

# 表碼：統計視角、當月欄、累計欄、預算欄
TABLES = {
    "6820": ("全國稅目", "C", "G", "L"),
    "6900": ("政府歸屬", "B", "F", "K"),
    "6920": ("中央稅目", "B", "F", "K"),
    "6951": ("地方稅目（含統籌）", "C", "G", "L"),
}

OWNERS = (
    "中央政府",
    "地方政府(含中央統籌)",
    "撥入中央特種基金",
)

COLUMNS = [
    "資料鍵",
    "年度",
    "月份",
    "來源表",
    "統計視角",
    "來源列",
    "項目",
    "原始項目",
    "項目型態",
    "當月實徵_億元",
    "累計實徵_億元",
    "全年預算_億元",
    "新聞網址",
    "Excel網址",
]


# ======================================
# 1. 讀取 XLSX 儲存格
# ======================================

def read_cells(filename):

    filename = Path(filename)

    if (
        not filename.is_file()
        or not zipfile.is_zipfile(filename)
    ):
        raise ValueError(
            f"Excel 不存在或格式無效：{filename}"
        )

    with zipfile.ZipFile(filename) as z:

        shared = []

        if "xl/sharedStrings.xml" in z.namelist():

            node = ET.fromstring(
                z.read("xl/sharedStrings.xml")
            )

            shared = [
                "".join(
                    text.text or ""
                    for text in si.findall(
                        ".//x:t", NS
                    )
                )
                for si in node.findall("x:si", NS)
            ]

        first_sheet = (
            "xl/worksheets/sheet1.xml"
        )

        if first_sheet not in z.namelist():
            raise ValueError(
                f"{filename.name} 缺少第一張工作表"
            )

        root = ET.fromstring(
            z.read(first_sheet)
        )

        cells = {}

        for cell in root.findall(
            ".//x:sheetData/x:row/x:c", NS
        ):

            match = re.fullmatch(
                r"([A-Z]+)(\d+)",
                cell.get("r", "")
            )

            if not match:
                continue

            value = cell.find("x:v", NS)
            inline = cell.find("x:is", NS)

            if (
                value is not None
                and value.text is not None
            ):

                if cell.get("t") == "s":
                    text = shared[int(value.text)]
                else:
                    text = value.text

            elif inline is not None:

                text = "".join(
                    t.text or ""
                    for t in inline.findall(
                        ".//x:t", NS
                    )
                )

            else:
                continue

            cells[(
                int(match.group(2)),
                match.group(1)
            )] = text

        return cells


# ======================================
# 2. 清理項目名稱
# ======================================

def item_name(text):

    return re.sub(
        r"\s+",
        "",
        str(text or "")
    )


# ======================================
# 3. 金額轉換：百萬元 → 億元
# ======================================

def amount_yi(raw):

    if (
        raw is None
        or str(raw).strip()
        in ("", "-", "—", "…")
    ):
        return None

    try:

        number = Decimal(
            str(raw).strip().replace(",", "")
        )

        if not number.is_finite():
            raise ValueError(
                f"非有限數值：{raw}"
            )

        return (
            number / Decimal("100")
        ).quantize(
            Decimal("0.01")
        )

    except InvalidOperation:

        raise ValueError(
            f"無法解析金額：{raw!r}"
        ) from None


# ======================================
# 4. 處理單一月份
# ======================================

def parse_month(entry):

    year = int(entry["year"])
    month = int(entry["month"])

    ym = f"{year}{month:02d}"

    if entry.get("status") != "success":
        raise ValueError(
            f"{ym} 下載狀態不是 success"
        )

    files = entry.get("files", {})

    if set(files) != set(TABLES):
        raise ValueError(
            f"{ym} 四表不足或代碼不符："
            f"{sorted(files)}"
        )

    records = []
    totals = {}
    owners = {}

    # 分別整理四張 Excel
    for code, config in TABLES.items():

        view, mon_col, cum_col, budget_col = (
            config
        )

        info = files[code]
        path = Path(info["path"])

        expected = f"{code}_{ym}.xlsx"

        if path.name != expected:
            raise ValueError(
                f"{ym} 來源檔案命名不符："
                f"{path.name}"
            )

        values = read_cells(path)

        count = 0

        for row in range(7, 150):

            raw_name = str(
                values.get((row, "A"), "")
            )

            name = item_name(raw_name)

            if name.startswith("說明"):
                break

            if not name:
                continue

            monthly = amount_yi(
                values.get((row, mon_col))
            )

            cumulative = amount_yi(
                values.get((row, cum_col))
            )

            budget = amount_yi(
                values.get((row, budget_col))
            )

            if all(
                x is None
                for x in (
                    monthly,
                    cumulative,
                    budget
                )
            ):
                continue

            count += 1

            records.append({
                "資料鍵":
                    f"{ym}-{code}-{row}",
                "年度": year,
                "月份": month,
                "來源表": code,
                "統計視角": view,
                "來源列": row,
                "項目": name,
                "原始項目": raw_name,
                "項目型態": (
                    "官方總計"
                    if name == "總計"
                    else "明細或小計"
                ),
                "當月實徵_億元": monthly,
                "累計實徵_億元": cumulative,
                "全年預算_億元": budget,
                "新聞網址": entry["news_url"],
                "Excel網址": info["url"],
            })

            # 記錄官方總計
            if name == "總計":

                if (
                    monthly is None
                    or cumulative is None
                ):
                    raise ValueError(
                        f"{ym} {code} "
                        "總計金額缺漏"
                    )

                totals[code] = (
                    monthly,
                    cumulative
                )

            # 記錄政府歸屬三大類
            if (
                code == "6900"
                and name in OWNERS
            ):

                if (
                    monthly is None
                    or cumulative is None
                ):
                    raise ValueError(
                        f"{ym} 政府別金額缺漏："
                        f"{name}"
                    )

                owners[name] = (
                    monthly,
                    cumulative
                )

        if count < 5:
            raise ValueError(
                f"{ym} {code} "
                f"有效資料不足：僅 {count} 筆"
            )

    # ======================================
    # 5. 核對官方總計
    # ======================================

    if set(totals) != set(TABLES):
        raise ValueError(
            f"{ym} 缺少四表官方總計"
        )

    if set(owners) != set(OWNERS):
        raise ValueError(
            f"{ym} 缺少政府別三大分類"
        )

    # 0 = 當月，1 = 累計
    for index, label in (
        (0, "當月"),
        (1, "累計")
    ):

        national = totals["6820"][index]

        if (
            national
            != totals["6900"][index]
        ):
            raise ValueError(
                f"{ym} {label}全國總計："
                "6820 與 6900 不一致"
            )

        if (
            totals["6920"][index]
            != owners[OWNERS[0]][index]
        ):
            raise ValueError(
                f"{ym} {label}中央總計"
                "與 6900 不一致"
            )

        if (
            totals["6951"][index]
            != owners[OWNERS[1]][index]
        ):
            raise ValueError(
                f"{ym} {label}地方總計"
                "與 6900 不一致"
            )

        gov_sum = sum(
            owners[name][index]
            for name in OWNERS
        )

        # 允許官方四捨五入尾差
        if abs(
            gov_sum - national
        ) > Decimal("0.03"):

            raise ValueError(
                f"{ym} {label}"
                "中央＋地方＋基金"
                "與全國差異過大"
            )

    # ======================================
    # 6. 產生檢核紀錄
    # ======================================

    check = {
        "year": year,
        "month": month,
        "record_count": len(records),
        "national_monthly_yi":
            str(totals["6900"][0]),
        "national_cumulative_yi":
            str(totals["6900"][1]),
        "central_cumulative_yi":
            str(owners[OWNERS[0]][1]),
        "local_cumulative_yi":
            str(owners[OWNERS[1]][1]),
        "fund_cumulative_yi":
            str(owners[OWNERS[2]][1]),
        "status": "passed",
    }

    return records, check


# ======================================
# 7. CSV 輸出
# ======================================

def write_csv(path, rows):

    tmp = path.with_name(
        path.name + ".tmp"
    )

    with tmp.open(
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=COLUMNS
        )

        writer.writeheader()

        for record in rows:

            writer.writerow({
                column: (
                    ""
                    if record[column] is None
                    else str(record[column])
                )
                for column in COLUMNS
            })

    tmp.replace(path)


# ======================================
# 8. 全年度整合
# ======================================

def run(
    manifest_path="data/manifest.json",
    output_dir="data"
):

    manifest_path = Path(
        manifest_path
    )

    output_dir = Path(
        output_dir
    )

    manifest = json.loads(
        manifest_path.read_text(
            encoding="utf-8"
        )
    )

    year = int(manifest["year"])

    entries = manifest.get(
        "months", []
    )

    if not entries:
        raise ValueError(
            "manifest.json 沒有月份資料"
        )

    months = [
        int(entry["month"])
        for entry in entries
    ]

    if len(months) != len(set(months)):
        raise ValueError(
            "manifest.json 出現重複月份"
        )

    if any(
        int(entry["year"]) != year
        for entry in entries
    ):
        raise ValueError(
            "manifest.json 混入其他年度"
        )

    if sorted(months) != list(
        range(1, max(months) + 1)
    ):
        raise ValueError(
            f"該年度月份不連續："
            f"{sorted(months)}"
        )

    all_rows = []
    checks = []

    for entry in sorted(
        entries,
        key=lambda e: int(e["month"])
    ):

        records, check = parse_month(
            entry
        )

        all_rows.extend(records)
        checks.append(check)

        print(
            f"✓ {year} 年 "
            f"{check['month']:02d} 月："
            f"{check['record_count']} 筆，"
            "全國累計 "
            f"{check['national_cumulative_yi']}"
            " 億元"
        )

    # ======================================
    # 9. 防止重複資料
    # ======================================

    keys = [
        row["資料鍵"]
        for row in all_rows
    ]

    if len(keys) != len(set(keys)):
        raise ValueError(
            "資料鍵重複；停止匯出"
        )

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    # ======================================
    # 10. 輸出統一資料
    # ======================================

    csv_path = (
        output_dir / "tax_records.csv"
    )

    write_csv(
        csv_path,
        all_rows
    )

    # ======================================
    # 11. 輸出檢核報告
    # ======================================

    validation = {
        "year": year,
        "months": checks,
        "total_records": len(all_rows),
        "status": "passed",
    }

    validation_path = (
        output_dir / "validation.json"
    )

    tmp = validation_path.with_name(
        validation_path.name + ".tmp"
    )

    tmp.write_text(
        json.dumps(
            validation,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    tmp.replace(
        validation_path
    )

    print(
        f"\n全部驗證通過："
        f"{len(checks)} 個月份，"
        f"共 {len(all_rows)} 筆"
    )

    print(
        f"CSV：{csv_path}"
    )

    print(
        f"檢核報告：{validation_path}"
    )

    return all_rows


# ======================================
# 12. GitHub Actions 執行入口
# ======================================

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description=(
            "驗證並整理財政部"
            "四張月度 Excel"
        )
    )

    parser.add_argument(
        "--manifest",
        default="data/manifest.json"
    )

    parser.add_argument(
        "--output",
        default="data"
    )

    args = parser.parse_args()

    run(
        manifest_path=args.manifest,
        output_dir=args.output
    )
