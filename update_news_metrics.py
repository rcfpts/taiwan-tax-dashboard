
"""同步財政部新聞稿月度指標至 Google Sheets。

僅更新「新聞稿月度指標」，不修改「稅收原始資料」。
修正：允許完全空白的工作表進行首次初始化。
"""

import csv
import json
from decimal import Decimal, InvalidOperation
from pathlib import Path

from news_metrics import COLUMNS, validate_metrics
from update_sheets import open_spreadsheet, column_letter


WORKSHEET_NAME = "新聞稿月度指標"

TEXT_COLUMNS = {
    "資料鍵", "資料類型", "預算版本",
    "新聞網址", "資料狀態", "註記",
}
INT_COLUMNS = {"年度", "月份"}

CHUNK_ROWS = 1000


def read_verified_metrics(
    csv_path="data/news_metrics.csv",
    report_path="data/news_metrics_validation.json",
):
    """只讀取已通過驗證的新聞稿指標。"""

    report = json.loads(
        Path(report_path).read_text(encoding="utf-8")
    )

    if report.get("status") != "passed":
        raise ValueError("新聞稿月度指標未通過驗證，停止同步")

    with Path(csv_path).open(
        "r", encoding="utf-8-sig", newline=""
    ) as f:
        reader = csv.DictReader(f)

        if reader.fieldnames != COLUMNS:
            raise ValueError("新聞稿指標 CSV 欄位格式不符")

        rows = list(reader)

    expected_keys = report.get("months", [])
    actual_keys = [r["資料鍵"] for r in rows]

    if (
        len(rows) != report.get("record_count")
        or actual_keys != expected_keys
    ):
        raise ValueError(
            "新聞稿指標 CSV 與檢核報告不一致"
        )

    validate_metrics(rows)
    return rows


def to_sheet_row(record):
    """將 CSV 資料轉成 Google Sheets 可寫入的型態。"""

    values = []

    for col in COLUMNS:
        value = record.get(col, "")

        if value is None or str(value) == "":
            values.append("")

        elif col in TEXT_COLUMNS:
            values.append(str(value))

        elif col in INT_COLUMNS:
            values.append(int(value))

        else:
            try:
                d = Decimal(str(value))

                if not d.is_finite():
                    raise ValueError(
                        f"數字無效：{col}"
                    )

                values.append(float(d))

            except InvalidOperation as exc:
                raise ValueError(
                    f"無法轉換數值：{col}={value!r}"
                ) from exc

    return values


def existing_rows(worksheet):
    """讀取既有資料，並判斷是否可以安全更新。

    完全空白：允許初始化。
    表頭正確：保留歷史紀錄。
    表頭錯誤且有內容：拒絕覆蓋。
    """

    rows = worksheet.get_all_values(
        value_render_option="UNFORMATTED_VALUE"
    )

    # 情況一：Google Sheets 完全沒有回傳資料
    if not rows:
        return []

    # 情況二：存在空白列，但整張表沒有任何內容
    if not any(
        value is not None and str(value).strip()
        for row in rows
        for value in row
    ):
        return []

    # 情況三：已經有資料，必須驗證表頭
    header = list(rows[0])

    while header and not str(header[-1]).strip():
        header.pop()

    if header != COLUMNS:
        raise ValueError(
            "既有『新聞稿月度指標』表頭不同，"
            "停止更新以避免覆蓋資料"
        )

    found = []
    keys = set()

    for idx, vals in enumerate(rows[1:], start=2):

        if not any(str(v).strip() for v in vals):
            continue

        padded = (
            list(vals) + [""] * len(COLUMNS)
        )[:len(COLUMNS)]

        record = dict(zip(COLUMNS, padded))

        key = str(record["資料鍵"])

        if not key or key in keys:
            raise ValueError(
                f"第 {idx} 列資料鍵無效或重複"
            )

        try:
            year = int(record["年度"])
            month = int(record["月份"])
        except (ValueError, TypeError) as exc:
            raise ValueError(
                f"第 {idx} 列年月格式錯誤"
            ) from exc

        if key != f"{year}{month:02d}":
            raise ValueError(
                f"第 {idx} 列資料鍵與年月不符"
            )

        keys.add(key)
        found.append(record)

    return found


def sync(worksheet, incoming):
    """更新收到的月份，保留其他歷史年度。"""

    validate_metrics(incoming)

    # 確保空白工作表可以通過此步驟
    existing = existing_rows(worksheet)

    incoming_keys = {
        str(r["資料鍵"]) for r in incoming
    }

    # 舊資料只移除這次要重新寫入的月份
    retained = [
        row for row in existing
        if str(row["資料鍵"]) not in incoming_keys
    ]

    merged = retained + incoming

    merged_keys = [
        str(r["資料鍵"]) for r in merged
    ]

    if len(merged_keys) != len(set(merged_keys)):
        raise ValueError("合併後出現重複資料鍵")

    merged.sort(
        key=lambda r: (
            int(r["年度"]),
            int(r["月份"]),
        )
    )

    # 寫入前先完成資料轉換
    values = [COLUMNS] + [
        to_sheet_row(r) for r in merged
    ]

    need_rows = len(values) + 20

    if (
        worksheet.row_count < need_rows
        or worksheet.col_count < len(COLUMNS)
    ):
        worksheet.resize(
            rows=max(worksheet.row_count, need_rows),
            cols=max(worksheet.col_count, len(COLUMNS)),
        )

    # 分批寫入
    for start in range(0, len(values), CHUNK_ROWS):
        worksheet.update(
            range_name=f"A{start + 1}",
            values=values[start:start + CHUNK_ROWS],
            value_input_option="RAW",
        )

    # 清除多餘的舊資料列
    worksheet.batch_clear([
        f"A{len(values) + 1}:"
        f"{column_letter(len(COLUMNS))}"
        f"{worksheet.row_count}"
    ])

    worksheet.freeze(rows=1)

    print(
        f"新聞稿月度指標："
        f"本次同步 {len(incoming)} 月，"
        f"保留歷史 {len(retained)} 月，"
        f"總計 {len(merged)} 月"
    )

    return len(merged)


def main():
    """執行正式 Google Sheets 同步。"""

    rows = read_verified_metrics()

    spreadsheet = open_spreadsheet()

    from gspread.exceptions import WorksheetNotFound

    try:
        worksheet = spreadsheet.worksheet(
            WORKSHEET_NAME
        )
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=100,
            cols=len(COLUMNS),
        )

    total = sync(worksheet, rows)

    print(
        f"完成寫入「{WORKSHEET_NAME}」："
        f"{total} 個月份"
    )


if __name__ == "__main__":
    main()
