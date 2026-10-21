"""財政部月度賦稅 Excel：解析、驗證、輸出長表。

整體流程（每次都從所有 Excel 重新計算，不讀回舊資料，所以可重複執行）：
    python crawler.py                 # 下載 → data/manifest.json、data/excel/
    python processor.py               # 解析、驗證 → data/tax_records.csv、data/validation.json
    python update_sheets.py           # 僅替換收到的年月，保留其他歷史資料

手動備援（財政部擋 GitHub 主機時）：
    把四張 Excel 放進 data/inbox/，然後
    python processor.py --from-dir data/inbox

輸出長表每一列 ＝ 某年月、某張表、某個稅目。只存當月、累計、全年預算三個核心金額；其餘分析留給儀表板。
"""

import argparse
import csv
import json
import re
import sys
import zipfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# 表碼：(統計視角, 政府別)；政府別為 None 者由第一層項目決定
TABLES = {
    "6820": ("全國稅目", "全國"),
    "6900": ("政府歸屬", None),
    "6920": ("中央稅目", "中央"),
    "6951": ("地方稅目（含統籌）", "地方"),
}

GOV_OF_NAME = {
    "中央政府": "中央",
    "地方政府(含中央統籌)": "地方",
    "撥入中央特種基金": "基金",
}

# 「其中項」：金額已包含在其他稅目內，不可再與其他稅目相加。
# 房地合一課徵所得稅已含在營利事業所得稅與綜合所得稅內。
# 若財政部改變表格結構，下方的「可加總列合計 ＝ 總計」檢查會失敗並停止。
MEMO_ITEMS = {("6820", "房地合一課徵所得稅")}

# 加總檢查容許的四捨五入誤差（百萬元）
TOLERANCE_MILLION = Decimal("5")

# 長表只保留三個跨月穩定的金額欄位（百萬元換算為億元）；
# 其餘均為資料識別、父子結構、是否可加總與來源追溯所需的中繼資料。
COLUMNS = [
    "資料鍵", "年度", "月份", "來源表", "視角", "政府別",
    "稅目", "稅目路徑", "父項目", "層級", "列型態", "計入加總",
    "當月實徵_億元", "累計實徵_億元", "全年預算_億元",
    "預算版本", "來源列", "新聞網址", "Excel網址",
]

# 財政部固定四表的核心儲存格位置，字母 A=0。
# 114/12 的年度表頭不同，但這三個欄位位置相同。
CORE_COLS = {
    "6820": {"month": 2, "cum": 6, "budget": 11},
    "6900": {"month": 1, "cum": 5, "budget": 10},
    "6920": {"month": 1, "cum": 5, "budget": 10},
    "6951": {"month": 2, "cum": 6, "budget": 11},
}


# ======================================
# 1. 讀取 XLSX（只用標準函式庫）
# ======================================

def _si_text(si):
    parts = [t.text or "" for t in si.findall("x:t", NS)]
    for run in si.findall("x:r", NS):
        parts += [t.text or "" for t in run.findall("x:t", NS)]
    return "".join(parts)


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_cells(filename):
    """回傳 {(列號, 欄索引從0起): 文字}。"""

    filename = Path(filename)

    if not filename.is_file() or not zipfile.is_zipfile(filename):
        raise ValueError(f"Excel 不存在或格式無效：{filename}")

    with zipfile.ZipFile(filename) as z:
        names = z.namelist()

        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [_si_text(si) for si in root.findall("x:si", NS)]

        sheets = sorted(
            n for n in names
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)
        )
        if not sheets:
            raise ValueError(f"{filename.name} 沒有工作表")

        root = ET.fromstring(z.read(sheets[0]))

        cells = {}
        for cell in root.findall(".//x:sheetData/x:row/x:c", NS):
            match = re.fullmatch(r"([A-Z]+)(\d+)", cell.get("r", ""))
            if not match:
                continue

            kind = cell.get("t")
            value = cell.find("x:v", NS)
            inline = cell.find("x:is", NS)

            if kind == "inlineStr" and inline is not None:
                text = _si_text(inline)
            elif value is not None and value.text is not None:
                text = shared[int(value.text)] if kind == "s" else value.text
            else:
                continue

            cells[(int(match.group(2)), _col_index(match.group(1)))] = text

        return cells


# ======================================
# 2. 名稱與數值清理
# ======================================

_FOOTNOTE = re.compile(r"\(\d+\)$")


def clean_name(text):
    """去掉所有空白（含全形）、統一括號、移除結尾腳註符號 (2)。"""

    s = re.sub(r"\s+", "", str(text or ""))
    s = s.replace("（", "(").replace("）", ")")
    return _FOOTNOTE.sub("", s)


def parse_amount(raw):
    """回傳 Decimal 或 None。財政部以 -- 表示無資料。"""

    if raw is None:
        return None

    text = str(raw).strip().replace(",", "")

    if text == "" or re.fullmatch(r"[-–—－…]+", text):
        return None

    try:
        number = Decimal(text)
    except InvalidOperation:
        raise ValueError(f"無法解析數值：{raw!r}") from None

    if not number.is_finite():
        raise ValueError(f"非有限數值：{raw!r}")

    return number


def to_yi(million):
    """百萬元 → 億元（精確到 0.01，無損）。"""

    if million is None:
        return ""
    return str((million / Decimal(100)).quantize(Decimal("0.01")))



# ======================================
# 3. 表頭定位（版面改變時明確報錯）
# ======================================

def _header_text(cells, col):
    return "".join(
        re.sub(r"\s+", "", cells.get((r, col), ""))
        for r in range(0, 7)
    )


def locate_columns(cells, label, code):
    """只定位四張表每月一致的當月、累計和全年預算。

    114 年 12 月的欄名「本年度實徵淨額」視為年末累計；
    其餘月份的「本年度累計實徵淨額」也是同一個累計欄位。
    只對三個核心表頭檢查，不要求年增率、分配預算等附屬欄位。
    """
    cols = CORE_COLS[code]
    labels = {
        "month": ("本月實徵淨額",),
        "cum": ("本年度累計實徵淨額", "本年度實徵淨額"),
        "budget": ("本年度預算數",),
    }
    for key, col in cols.items():
        text = _header_text(cells, col)
        if not any(expected in text for expected in labels[key]):
            raise ValueError(
                f"{label} 核心欄 {key}（Excel 第 {col + 1} 欄）表頭不符：{text!r}；"
                "暫停匯入以免抓錯金額"
            )
    return cols.copy()


def find_period(cells):
    for (row, _col), value in sorted(cells.items()):
        if row > 4:
            continue
        match = re.search(r"(\d{2,3})年\s*(\d{1,2})月", value)
        if match:
            return int(match.group(1)), int(match.group(2))
    return None


# ======================================
# 4. 解析單一張表
# ======================================

def _indent(raw):
    raw = str(raw)
    return len(raw) - len(raw.lstrip(" \t　"))


def _is_decimal_or_missing(value):
    if value is None:
        return True
    v = str(value).strip().replace(",", "")
    if not v or re.fullmatch(r"[-–—－…]+", v):
        return True
    try:
        return Decimal(v).is_finite()
    except InvalidOperation:
        return False


def parse_table(code, path, year, month):
    label = f"{year}{month:02d} 表{code}"
    cells = read_cells(path)

    period = find_period(cells)
    if period != (year, month):
        raise ValueError(f"{label} 檔案內的期間是 {period}，與預期不符")

    cols = locate_columns(cells, label, code)

    by_row = defaultdict(dict)
    for (row, col), value in cells.items():
        by_row[row][col] = value

    start = next(
        (r for r in sorted(by_row)
         if clean_name(by_row[r].get(0, "")) == "總計"),
        None,
    )
    if start is None:
        raise ValueError(f"{label} 找不到「總計」列")

    raw_rows = []
    notes = []
    in_notes = False

    for row in sorted(r for r in by_row if r >= start):
        raw_label = str(by_row[row].get(0, ""))
        name = clean_name(raw_label)

        # 官方說明可能在表格下方拆成多行，其中 C/D 欄會出現
        # 「百萬元，累計1-1月實物抵繳金額共為」之類文字，不能當金額。
        # 優先判斷整列是否為附註，而非嘗試把註解欄位轉成 Decimal。
        if (
            name.startswith(("說明", "備註", "附註", "資料來源"))
            or re.match(r"^(?:\(?\d+\)?[.、：:]|[註注][：:])", name)
        ):
            in_notes = True

        if in_notes:
            notes.extend(str(v) for v in by_row[row].values())
            continue

        if not name:
            continue

        # 欄位中含文字但整列又不像稅目：
        # 114 年 1 月附表可能把跨欄說明拆入數值欄。
        # 僅跳過可明確判別為註記的列；真正稅目數字格式錯誤仍會報錯。
        numeric_values = [
            by_row[row].get(col) for col in cols.values() if col is not None
        ]
        note_words = ("百萬元", "實物抵繳", "共為", "撥入金融業特別準備金施行")
        if (
            len(raw_rows) >= 5
            and (any(w in name for w in note_words)
                 or any(
                     isinstance(v, str)
                     and any(w in v for w in note_words)
                     and not _is_decimal_or_missing(v)
                     for v in numeric_values
                 ))
        ):
            in_notes = True
            notes.extend(str(v) for v in by_row[row].values())
            continue

        amounts = {
            key: parse_amount(by_row[row].get(col)) if col is not None else None
            for key, col in cols.items()
        }

        if all(amounts[k] is None for k in ("month", "cum", "budget")):
            continue

        raw_rows.append({
            "src_row": row,
            "name": name,
            "indent": _indent(raw_label),
            **amounts,
        })

    if not raw_rows or raw_rows[0]["name"] != "總計":
        raise ValueError(f"{label} 第一筆資料不是總計")

    rows = _build_hierarchy(code, raw_rows)

    paths = [r["path"] for r in rows]
    if len(paths) != len(set(paths)):
        raise ValueError(f"{label} 稅目路徑重複")

    return rows, "\n".join(notes)


def _build_hierarchy(code, raw_rows):
    """以官方排版縮排判斷層級，並標出哪些列可以加總。"""

    total, items = raw_rows[0], raw_rows[1:]
    view, fixed_gov = TABLES[code]

    total.update(
        level=0, parent=None, path="總計", parent_name="",
        has_children=False, kind="總計", include=False, gov="全國",
    )

    if not items:
        return [total]

    base = min(r["indent"] for r in items)
    stack = []

    for rec in items:
        level = rec["indent"] - base + 1

        while stack and stack[-1]["level"] >= level:
            stack.pop()

        parent = stack[-1] if stack else None
        rec.update(
            level=level,
            parent=parent,
            parent_name=parent["name"] if parent else "",
            path=(parent["path"] + "/" if parent else "") + rec["name"],
            has_children=False,
        )
        if parent:
            parent["has_children"] = True
        stack.append(rec)

    for rec in items:
        top = rec
        while top["parent"]:
            top = top["parent"]

        rec["gov"] = fixed_gov or GOV_OF_NAME.get(top["name"], "")

        memo = False
        node = rec
        while node:
            if (code, node["name"]) in MEMO_ITEMS:
                memo = True
            node = node["parent"]

        own_memo = (code, rec["name"]) in MEMO_ITEMS

        if own_memo:
            rec["kind"] = "其中項"
        elif memo:
            rec["kind"] = "其中項明細"
        elif rec["has_children"]:
            rec["kind"] = "小計"
        else:
            rec["kind"] = "明細"

        rec["include"] = rec["kind"] == "明細"

    return [total] + items


# ======================================
# 5. 單月驗證
# ======================================

def _close(a, b):
    return a is not None and b is not None and abs(a - b) <= TOLERANCE_MILLION


def check_month(year, month, tables):
    """回傳 (check, warnings)；發現錯誤時丟出 ValueError。"""

    ym = f"{year}{month:02d}"
    problems = []
    warnings = []

    def total(code):
        return tables[code][0]

    def owner(name):
        return next(
            r for r in tables["6900"]
            if r["level"] == 1 and r["name"] == name
        )

    try:
        owners = {name: owner(name) for name in GOV_OF_NAME}
    except StopIteration:
        raise ValueError(f"{ym} 6900 缺少中央／地方／基金三大類") from None

    for key, label in (("month", "當月"), ("cum", "累計")):

        national = total("6820")[key]

        if national != total("6900")[key]:
            problems.append(f"{label}全國總計：表 6820 與 6900 不一致")

        if total("6920")[key] != owners["中央政府"][key]:
            problems.append(f"{label}中央總計：表 6920 與 6900 不一致")

        if total("6951")[key] != owners["地方政府(含中央統籌)"][key]:
            problems.append(f"{label}地方總計：表 6951 與 6900 不一致")

        gov_sum = sum(o[key] for o in owners.values())
        if not _close(gov_sum, national):
            problems.append(
                f"{label}中央＋地方＋基金 {gov_sum} 與全國 {national} 差異過大"
            )

        # 可加總列合計必須等於各表總計（抓出重複計算或漏列）
        for code, rows in tables.items():
            leaf_sum = sum(
                (r[key] or Decimal(0)) for r in rows if r["include"]
            )
            if not _close(leaf_sum, rows[0][key]):
                problems.append(
                    f"表 {code} {label}：可加總列合計 {leaf_sum} "
                    f"≠ 總計 {rows[0][key]}（百萬元）"
                )

    for code, rows in tables.items():
        leaf_budget = sum(
            (r["budget"] or Decimal(0)) for r in rows if r["include"]
        )
        if not _close(leaf_budget, rows[0]["budget"]):
            warnings.append(
                f"{ym} 表 {code} 全年預算：可加總列合計 {leaf_budget} "
                f"≠ 總計 {rows[0]['budget']}（百萬元）"
            )

    if problems:
        raise ValueError(f"{ym} 驗證失敗：" + "；".join(problems))

    check = {
        "year": year,
        "month": month,
        "national_monthly_yi": to_yi(total("6900")["month"]),
        "national_cumulative_yi": to_yi(total("6900")["cum"]),
        "central_cumulative_yi": to_yi(owners["中央政府"]["cum"]),
        "local_cumulative_yi": to_yi(owners["地方政府(含中央統籌)"]["cum"]),
        "fund_cumulative_yi": to_yi(owners["撥入中央特種基金"]["cum"]),
        "status": "passed",
    }

    return check, warnings


# ======================================
# 6. 單月 → 輸出列
# ======================================

def budget_version(notes):
    return "依法定數" if "依法定數" in notes else "原預算數"


def parse_month(entry):
    year = int(entry["year"])
    month = int(entry["month"])
    ym = f"{year}{month:02d}"

    files = entry.get("files", {})
    if set(files) != set(TABLES):
        raise ValueError(f"{ym} 四表不足或代碼不符：{sorted(files)}")

    tables = {}
    notes = []

    for code in TABLES:
        rows, note = parse_table(code, files[code]["path"], year, month)
        tables[code] = rows
        notes.append(note)

    check, warnings = check_month(year, month, tables)

    version = budget_version("\n".join(notes))

    records = []

    for code, rows in tables.items():
        view, _ = TABLES[code]

        for rec in rows:
            out = {
                "資料鍵": f"{ym}|{code}|{rec['path']}",
                "年度": year,
                "月份": month,
                "來源表": code,
                "視角": view,
                "政府別": rec["gov"],
                "稅目": rec["name"],
                "稅目路徑": rec["path"],
                "父項目": rec["parent_name"],
                "層級": rec["level"],
                "列型態": rec["kind"],
                "計入加總": "TRUE" if rec["include"] else "FALSE",
                "預算版本": version,
                "來源列": rec["src_row"],
                "新聞網址": entry.get("news_url", ""),
                "Excel網址": files[code].get("url", ""),
            }
            out["當月實徵_億元"] = to_yi(rec["month"])
            out["累計實徵_億元"] = to_yi(rec["cum"])
            out["全年預算_億元"] = to_yi(rec["budget"])
            records.append(out)

    check["record_count"] = len(records)
    check["budget_version"] = version

    return records, check, warnings


# ======================================
# 7. 輸入來源：manifest 或資料夾
# ======================================

_FILE_PATTERN = re.compile(r"(\d{4})_(\d{3})(\d{2})")


def entries_from_dir(folder):
    """手動備援：掃描資料夾，依檔名中的「表碼_民國年月」分組。"""

    groups = defaultdict(dict)

    for path in sorted(Path(folder).rglob("*.xlsx")):
        match = _FILE_PATTERN.search(path.name)
        if not match or match.group(1) not in TABLES:
            continue
        code, year, month = match.group(1), int(match.group(2)), int(match.group(3))
        groups[(year, month)][code] = {"path": str(path), "url": ""}

    return [
        {
            "year": y, "month": m, "news_url": "",
            "files": files, "status": "success",
        }
        for (y, m), files in sorted(groups.items())
    ]


def load_entries(manifest_path, from_dir):
    if from_dir:
        entries = entries_from_dir(from_dir)
        if not entries:
            raise ValueError(f"{from_dir} 裡找不到符合 表碼_年月.xlsx 的檔案")
        return entries

    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    entries = manifest.get("months", [])
    if not entries:
        raise ValueError("manifest.json 沒有月份資料")
    return entries


# ======================================
# 8. 輸出
# ======================================

def write_csv(path, rows):
    tmp = path.with_name(path.name + ".tmp")

    with tmp.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows({c: r.get(c, "") for c in COLUMNS} for r in rows)

    tmp.replace(path)


def write_json(path, data):
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    tmp.replace(path)


def run(manifest_path="data/manifest.json", output_dir="data", from_dir=None):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    entries = load_entries(manifest_path, from_dir)
    entries.sort(key=lambda e: (int(e["year"]), int(e["month"])))

    seen = [(int(e["year"]), int(e["month"])) for e in entries]
    if len(seen) != len(set(seen)):
        raise ValueError("來源出現重複的年月")

    all_rows, checks, warnings, errors = [], [], [], []

    for entry in entries:
        ym = f"{entry['year']}{int(entry['month']):02d}"

        if entry.get("status") != "success":
            warnings.append(f"{ym} 下載失敗，已略過：{entry.get('error', '')}")
            continue

        try:
            records, check, warn = parse_month(entry)
        except ValueError as exc:
            errors.append(str(exc))
            continue

        all_rows.extend(records)
        checks.append(check)
        warnings.extend(warn)

        print(
            f"✓ {entry['year']} 年 {int(entry['month']):02d} 月："
            f"{check['record_count']} 筆，全國累計 "
            f"{check['national_cumulative_yi']} 億元（預算：{check['budget_version']}）"
        )

    # 缺月只警告，不阻擋其他月份
    by_year = defaultdict(set)
    for c in checks:
        by_year[c["year"]].add(c["month"])
    latest = max(by_year) if by_year else None
    for year, months in sorted(by_year.items()):
        expected = max(months) if year == latest else 12
        missing = sorted(set(range(1, expected + 1)) - months)
        if missing:
            warnings.append(f"民國 {year} 年缺少月份：{missing}")

    keys = [r["資料鍵"] for r in all_rows]
    if len(keys) != len(set(keys)):
        errors.append("資料鍵重複")

    status = "failed" if errors else (
        "passed_with_warnings" if warnings else "passed"
    )

    validation = {
        "status": status,
        "years": sorted(by_year),
        "months": checks,
        "total_records": len(all_rows),
        "warnings": warnings,
        "errors": errors,
    }
    write_json(output_dir / "validation.json", validation)

    for text in warnings:
        print("⚠", text)

    if errors:
        for text in errors:
            print("✗", text)
        raise ValueError(f"驗證失敗，共 {len(errors)} 項；未輸出 CSV")

    write_csv(output_dir / "tax_records.csv", all_rows)

    print(f"\n完成：{len(checks)} 個月份，共 {len(all_rows)} 筆（{status}）")
    return all_rows


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="解析並驗證財政部月度賦稅 Excel")
    parser.add_argument("--manifest", default="data/manifest.json")
    parser.add_argument("--output", default="data")
    parser.add_argument(
        "--from-dir", default=None,
        help="手動備援：直接掃描此資料夾內的 Excel，不使用 manifest",
    )
    args = parser.parse_args()

    try:
        run(args.manifest, args.output, args.from_dir)
    except ValueError as exc:
        print(f"錯誤：{exc}", file=sys.stderr)
        sys.exit(1)
