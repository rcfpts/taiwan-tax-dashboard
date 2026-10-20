
"""將財政部驗證完成的稅收資料同步到 Google 試算表。"""

import csv
import json
import os

from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path

from processor import COLUMNS


# ======================================
# 1. Google 試算表設定
# ======================================

SPREADSHEET_ID = (
    "1TvYwDgFbCR4oV8qJnFKfHOFYNebsimbLlN9UUlXr8es"
)

WORKSHEET_NAME = "稅收原始資料"

AMOUNT_COLUMNS = (
    "當月實徵_億元",
    "累計實徵_億元",
    "全年預算_億元",
)

INTEGER_COLUMNS = (
    "年度",
    "月份",
    "來源列",
)

REQUIRED_CODES = {
    "6820",
    "6900",
    "6920",
    "6951",
}


# ======================================
# 2. 讀取已驗證資料
# ======================================

def load_verified_data(
    csv_path="data/tax_records.csv",
    validation_path="data/validation.json",
):

    report = json.loads(
        Path(validation_path).read_text(
            encoding="utf-8"
        )
    )

    if report.get("status") != "passed":
        raise ValueError(
            "資料尚未通過驗證，停止同步"
        )

    with Path(csv_path).open(
        encoding="utf-8-sig",
        newline=""
    ) as file:

        reader = csv.DictReader(file)

        if reader.fieldnames != COLUMNS:
            raise ValueError(
                "CSV 欄位格式不正確"
            )

        rows = list(reader)

    if not rows:
        raise ValueError("CSV 沒有資料")

    year = int(report["year"])

    checks = report.get("months", [])

    checks_by_month = {
        int(item["month"]): item
        for item in checks
    }

    if (
        not checks
        or len(checks_by_month) != len(checks)
    ):
        raise ValueError(
            "月份檢核紀錄有缺漏或重複"
        )

    expected_months = list(
        range(
            1,
            max(checks_by_month) + 1
        )
    )

    if sorted(checks_by_month) != expected_months:
        raise ValueError(
            "月份不連續，拒絕同步"
        )

    if any(
        item.get("status") != "passed"
        for item in checks
    ):
        raise ValueError(
            "部分月份未通過驗證"
        )

    keys = set()
    counts = Counter()
    codes = {}
    totals = {}

    for row in rows:

        y = int(row["年度"])
        m = int(row["月份"])

        code = row["來源表"]
        source_row = int(row["來源列"])

        key = (
            f"{y}{m:02d}-{code}-{source_row}"
        )

        if (
            y != year
            or m not in checks_by_month
        ):
            raise ValueError(
                "CSV 包含未檢核的年度或月份"
            )

        if (
            code not in REQUIRED_CODES
            or row["資料鍵"] != key
            or key in keys
        ):
            raise ValueError(
                f"資料鍵異常：{row['資料鍵']}"
            )

        keys.add(key)

        counts[m] += 1

        codes.setdefault(
            m, set()
        ).add(code)

        if (
            code == "6900"
            and row["項目"] == "總計"
        ):

            if m in totals:
                raise ValueError(
                    f"{m} 月官方總計重複"
                )

            totals[m] = Decimal(
                row["累計實徵_億元"]
            )

    if (
        int(report["total_records"])
        != len(rows)
    ):
        raise ValueError(
            "CSV 筆數與檢核報告不一致"
        )

    for month, check in checks_by_month.items():

        if (
            counts[month]
            != int(check["record_count"])
            or codes.get(month) != REQUIRED_CODES
        ):
            raise ValueError(
                f"{month} 月四表資料不完整"
            )

        if totals.get(month) != Decimal(
            check["national_cumulative_yi"]
        ):
            raise ValueError(
                f"{month} 月總額不一致"
            )

    return year, rows


# ======================================
# 3. 轉換 Google Sheets 儲存格格式
# ======================================

def as_sheet_row(record):

    result = []

    for field in COLUMNS:

        value = record.get(
            field, ""
        )

        if (
            value is None
            or str(value).strip() == ""
        ):
            result.append("")

        elif field in INTEGER_COLUMNS:

            result.append(
                int(value)
            )

        elif field in AMOUNT_COLUMNS:

            try:
                result.append(
                    float(
                        Decimal(str(value))
                    )
                )

            except InvalidOperation as exc:
                raise ValueError(
                    f"{field} 數值錯誤：{value}"
                ) from exc

        else:

            result.append(
                str(value)
            )

    return result


# ======================================
# 4. Google Service Account 驗證
# ======================================

def open_spreadsheet():

    raw_credentials = os.environ.get(
        "GOOGLE_SERVICE_ACCOUNT_JSON"
    )

    if not raw_credentials:
        raise RuntimeError(
            "缺少 GitHub Secret："
            "GOOGLE_SERVICE_ACCOUNT_JSON"
        )

    from google.oauth2.service_account import (
        Credentials
    )

    import gspread

    credentials_info = json.loads(
        raw_credentials
    )

    credentials = (
        Credentials.from_service_account_info(
            credentials_info,
            scopes=[
                "https://www.googleapis.com/"
                "auth/spreadsheets"
            ]
        )
    )

    client = gspread.authorize(
        credentials
    )

    sheet_id = os.environ.get(
        "GOOGLE_SHEET_ID",
        SPREADSHEET_ID
    )

    return client.open_by_key(
        sheet_id
    )


# ======================================
# 5. 建立或取得資料工作表
# ======================================

def get_or_create_worksheet(
    spreadsheet
):

    from gspread.exceptions import (
        WorksheetNotFound
    )

    try:

        return spreadsheet.worksheet(
            WORKSHEET_NAME
        )

    except WorksheetNotFound:

        return spreadsheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=1500,
            cols=len(COLUMNS)
        )


# ======================================
# 6. 同步指定年度資料
# ======================================

def sync_year(
    worksheet,
    year,
    records
):

    existing = (
        worksheet.get_all_values()
    )

    historical = []
    original_year_count = 0
    old_keys = set()

    if existing:

        if existing[0] != COLUMNS:
            raise ValueError(
                "工作表已有其他格式資料；"
                "為避免覆寫，停止更新"
            )

        for line_number, row in enumerate(
            existing[1:],
            start=2
        ):

            padded = (
                row + [""] * len(COLUMNS)
            )[:len(COLUMNS)]

            if not any(padded):
                continue

            key = padded[0].strip()

            try:

                existing_year = int(
                    padded[1]
                )

            except (
                ValueError,
                TypeError
            ) as exc:

                raise ValueError(
                    f"第 {line_number} 列"
                    "年度資料有誤"
                ) from exc

            if (
                not key
                or key in old_keys
            ):
                raise ValueError(
                    "既有資料鍵缺漏或重複"
                )

            old_keys.add(key)

            if existing_year == year:

                original_year_count += 1

            else:

                historical.append(
                    padded
                )

    # 只重建最新統計年度
    # 其他年度完整保留
    current = [
        as_sheet_row(item)
        for item in records
    ]

    all_values = (
        [COLUMNS]
        + historical
        + current
    )

    # 確保工作表容量足夠
    if (
        worksheet.row_count
        < len(all_values)
    ):

        worksheet.add_rows(
            len(all_values)
            - worksheet.row_count
        )

    if (
        worksheet.col_count
        < len(COLUMNS)
    ):

        worksheet.add_cols(
            len(COLUMNS)
            - worksheet.col_count
        )

    # 一次寫入所有資料
    # RAW 確保文字不會被當成公式
    worksheet.update(
        range_name=(
            f"A1:N{len(all_values)}"
        ),
        values=all_values,
        value_input_option="RAW"
    )

    # 如果舊資料比新資料多
    # 清除多餘尾列
    if len(existing) > len(all_values):

        worksheet.batch_clear([
            f"A{len(all_values)+1}:"
            f"N{len(existing)}"
        ])

    # 凍結標題列
    worksheet.freeze(
        rows=1
    )

    print(
        "Google Sheets 同步完成"
    )

    print(
        f"保留歷年資料："
        f"{len(historical)} 筆"
    )

    print(
        f"本年度原有："
        f"{original_year_count} 筆"
    )

    print(
        f"本年度更新後："
        f"{len(current)} 筆"
    )


# ======================================
# 7. 程式入口
# ======================================

def main():

    # 先確認資料通過驗證
    year, rows = load_verified_data()

    # 再連線 Google Sheets
    spreadsheet = open_spreadsheet()

    worksheet = get_or_create_worksheet(
        spreadsheet
    )

    sync_year(
        worksheet,
        year,
        rows
    )

    print(
        f"完成工作表：{WORKSHEET_NAME}"
    )

    print(
        "https://docs.google.com/"
        f"spreadsheets/d/{spreadsheet.id}/edit"
    )


if __name__ == "__main__":
    main()
