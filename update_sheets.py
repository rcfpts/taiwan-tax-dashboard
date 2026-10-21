"""同步通過驗證的稅收資料到 Google 試算表。

設計原則：
- 對本次實際抓到的年月進行全月替換，保留其餘歷史年月。
- 讀取既有工作表以避免跨年覆蓋掉更早年度；資料鍵不得重複。
- 發布前檢查最新統計年度從 1 月到最新月份是否連續。
- 不自動覆蓋未知表頭，舊版已存在資料必須先安排遷移。
- 優先完整檢核及安全拒絕，避免資料少抓時整表消失。
"""

import csv
import json
import os
from decimal import Decimal, InvalidOperation
from collections import Counter
from pathlib import Path

from processor import COLUMNS


SPREADSHEET_ID = "1TvYwDgFbCR4oV8qJnFKfHOFYNebsimbLlN9UUlXr8es"
WORKSHEET_NAME = "稅收原始資料"

# 舊版程式寫入的表頭；看到它代表是自己先前的資料，可以放心覆蓋
LEGACY_COLUMNS = [
    "資料鍵", "年度", "月份", "來源表", "統計視角", "來源列", "項目",
    "原始項目", "項目型態", "當月實徵_億元", "累計實徵_億元",
    "全年預算_億元", "新聞網址", "Excel網址",
]

INTEGER_COLUMNS = {"年度", "月份", "層級", "來源列"}
BOOLEAN_COLUMNS = {"計入加總"}
TEXT_COLUMNS = {
    "資料鍵", "來源表", "視角", "政府別", "稅目", "稅目路徑", "父項目",
    "列型態", "預算版本", "財劃法版本", "新聞網址", "Excel網址",
}

CHUNK_ROWS = 2000


# ======================================
# 1. 讀取並檢查已驗證資料
# ======================================

def load_verified_data(
    csv_path="data/tax_records.csv",
    validation_path="data/validation.json",
):
    report = json.loads(Path(validation_path).read_text(encoding="utf-8"))

    if report.get("status") not in ("passed", "passed_with_warnings"):
        raise ValueError("資料尚未通過驗證，停止同步")

    with Path(csv_path).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)

        if reader.fieldnames != COLUMNS:
            raise ValueError("CSV 欄位與資料字典不一致")

        rows = list(reader)

    if not rows:
        raise ValueError("CSV 沒有資料")

    # 防止 CSV 雖然筆數相同，卻混入不同月份或漏表。
    checks = report.get("months", [])
    expected_counts = Counter((int(x["year"]), int(x["month"])) for x in checks)
    if not expected_counts or any(n != 1 for n in expected_counts.values()):
        raise ValueError("檢核報告的年月資料缺漏或重複")
    if any(x.get("status") != "passed" for x in checks):
        raise ValueError("有月份尚未通過完整檢核")
    csv_counts = Counter((int(r["年度"]), int(r["月份"])) for r in rows)
    if set(expected_counts) != set(csv_counts):
        raise ValueError("CSV 與檢核報告的年月不一致")
    for check in checks:
        period = (int(check["year"]), int(check["month"]))
        if csv_counts[period] != int(check["record_count"]):
            raise ValueError(f"{period} 的資料筆數與檢核報告不一致")
        if {r["來源表"] for r in rows if (int(r["年度"]), int(r["月份"])) == period} != {"6820", "6900", "6920", "6951"}:
            raise ValueError(f"{period} 的四張附表不完整")

    if int(report["total_records"]) != len(rows):
        raise ValueError("CSV 筆數與檢核報告不一致")

    keys = [r["資料鍵"] for r in rows]
    if len(keys) != len(set(keys)):
        raise ValueError("資料鍵重複")

    return rows


# ======================================
# 2. 轉成試算表儲存格（一律用正確型態）
# ======================================

def to_cell(column, value):
    value = "" if value is None else str(value).strip()

    if value == "":
        return ""

    if column in INTEGER_COLUMNS:
        return int(value)

    if column in BOOLEAN_COLUMNS:
        if value.upper() not in ("TRUE", "FALSE"):
            raise ValueError(f"{column} 不是 TRUE/FALSE：{value}")
        return value.upper() == "TRUE"

    if column in TEXT_COLUMNS:
        return value

    try:
        return float(Decimal(value))
    except InvalidOperation as exc:
        raise ValueError(f"{column} 數值錯誤：{value}") from exc


def as_sheet_row(record):
    return [to_cell(column, record.get(column)) for column in COLUMNS]


# ======================================
# 3. 連線
# ======================================

def open_spreadsheet():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if not raw:
        raise RuntimeError("缺少 GitHub Secret：GOOGLE_SERVICE_ACCOUNT_JSON")

    import gspread
    from google.oauth2.service_account import Credentials

    credentials = Credentials.from_service_account_info(
        json.loads(raw),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )

    client = gspread.authorize(credentials)

    # 環境變數存在但是空字串時，也退回預設值
    sheet_id = os.environ.get("GOOGLE_SHEET_ID") or SPREADSHEET_ID

    return client.open_by_key(sheet_id)


def get_or_create_worksheet(spreadsheet):
    from gspread.exceptions import WorksheetNotFound

    try:
        return spreadsheet.worksheet(WORKSHEET_NAME)
    except WorksheetNotFound:
        return spreadsheet.add_worksheet(
            title=WORKSHEET_NAME, rows=2000, cols=len(COLUMNS)
        )


# ======================================
# 4. 整張覆蓋寫入
# ======================================

def column_letter(number):
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _existing_rows(worksheet):
    """讀取資料並檢查是否可安全沿用；不接受無法解析的歷史資料。"""
    existing = worksheet.get_all_values(value_render_option="UNFORMATTED_VALUE")
    if not existing:
        return []

    header = existing[0]
    while header and header[-1] == "":
        header.pop()

    if not header:
        if any(any(str(v).strip() for v in row) for row in existing[1:]):
            raise ValueError("試算表有資料但缺少表頭，拒絕覆蓋")
        return []

    if header == LEGACY_COLUMNS:
        if any(any(str(v).strip() for v in row) for row in existing[1:]):
            raise ValueError(
                "偵測到舊版 14 欄資料（舊版沒有層級與加總標記，內含重複計算的房地合一）。"
                "請先把該工作表複製一份備份，再清空原工作表整頁後重新執行"
            )
        return []

    if header != COLUMNS:
        raise ValueError("既有工作表表頭不是目前版本，為避免覆蓋已停止")

    records = []
    keys = set()
    for index, row in enumerate(existing[1:], start=2):
        padded = (row + [""] * len(COLUMNS))[:len(COLUMNS)]
        if not any(padded):
            continue
        record = dict(zip(COLUMNS, padded))
        key = str(record["資料鍵"]).strip()
        if not key or key in keys:
            raise ValueError(f"既有工作表第 {index} 列資料鍵缺漏或重複")
        keys.add(key)
        try:
            year, month = int(record["年度"]), int(record["月份"])
        except (ValueError, TypeError) as exc:
            raise ValueError(f"既有工作表第 {index} 列的年月無效") from exc
        if not (1 <= month <= 12):
            raise ValueError(f"既有工作表第 {index} 列月份不在 1~12")
        records.append(record)
    return records


def sync(worksheet, rows, require_current_contiguous=True):
    """整月更新收到的年月、保留其他歷年資料，不會誤刪舊年度。"""
    existing = _existing_rows(worksheet)

    incoming_periods = {(int(r["年度"]), int(r["月份"])) for r in rows}
    if not incoming_periods:
        raise ValueError("沒有可同步資料")

    # 單一年月可能修正稅目階層／名稱，因此要整月替換，而不是逐鍵合併。
    kept = [
        record for record in existing
        if (int(record["年度"]), int(record["月份"])) not in incoming_periods
    ]
    merged = kept + rows

    keys = [r["資料鍵"] for r in merged]
    if len(keys) != len(set(keys)):
        raise ValueError("合併後資料鍵重複，已停止寫入")

    # 尤其是最新年度，1～最新公布月份不可缺月；不完整就拒絕發布。
    if require_current_contiguous:
        latest = max(int(r["年度"]) for r in merged)
        months = {int(r["月份"]) for r in merged if int(r["年度"]) == latest}
        expected = set(range(1, max(months) + 1))
        if months != expected:
            raise ValueError(
                f"最新年度 {latest} 年仍缺少 {sorted(expected - months)} 月；"
                "為避免儀表板展示不完整歷史資料，停止同步"
            )

    # 轉型在觸碰工作表之前全部完成。
    merged.sort(key=lambda r: (
        int(r["年度"]), int(r["月份"]), str(r["來源表"]), str(r["稅目路徑"])
    ))
    values = [COLUMNS] + [as_sheet_row(r) for r in merged]
    last_col = column_letter(len(COLUMNS))

    needed_rows = len(values) + 20
    if worksheet.row_count < needed_rows or worksheet.col_count < len(COLUMNS):
        worksheet.resize(
            rows=max(worksheet.row_count, needed_rows),
            cols=max(worksheet.col_count, len(COLUMNS)),
        )

    # 先準備所有值，再分批寫入；寫入異常時 Actions 標示失敗。
    for start in range(0, len(values), CHUNK_ROWS):
        worksheet.update(
            range_name=f"A{start + 1}",
            values=values[start:start + CHUNK_ROWS],
            value_input_option="RAW",
        )

    worksheet.batch_clear([
        f"A{len(values) + 1}:{last_col}{worksheet.row_count}",
    ])
    if worksheet.col_count > len(COLUMNS):
        worksheet.batch_clear([
            f"{column_letter(len(COLUMNS) + 1)}1:"
            f"{column_letter(worksheet.col_count)}{worksheet.row_count}",
        ])

    worksheet.freeze(rows=1)
    print(f"本次更新 {len(rows)} 筆，保留其他年月 {len(kept)} 筆")
    return len(merged)


def contiguous_required():
    """預設要求最新年度 1 月到最新月份不可缺月。

    財政部若確定某個月不會補發，可設定環境變數 ALLOW_MISSING_MONTHS=1 放行，
    否則缺的那個月會讓所有更新永遠停住。
    """

    return os.environ.get("ALLOW_MISSING_MONTHS") != "1"


def main():
    rows = load_verified_data()

    spreadsheet = open_spreadsheet()
    worksheet = get_or_create_worksheet(spreadsheet)

    count = sync(worksheet, rows, require_current_contiguous=contiguous_required())

    print(f"完成：寫入 {count} 筆到「{WORKSHEET_NAME}」")
    print(f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}/edit")


if __name__ == "__main__":
    main()
