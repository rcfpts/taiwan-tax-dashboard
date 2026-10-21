"""僅新增／更新『新聞稿月度指標』，不更動已成功的『稅收原始資料』。"""

import csv
import json
from pathlib import Path
from decimal import Decimal, InvalidOperation

from news_metrics import COLUMNS, validate_metrics
from update_sheets import open_spreadsheet, column_letter

WORKSHEET_NAME = "新聞稿月度指標"
TEXT_COLUMNS = {"資料鍵", "資料類型", "預算版本", "新聞網址", "資料狀態", "註記"}
INT_COLUMNS = {"年度", "月份"}
CHUNK_ROWS = 1000


def read_verified_metrics(csv_path="data/news_metrics.csv", report_path="data/news_metrics_validation.json"):
    report = json.loads(Path(report_path).read_text(encoding="utf-8"))
    if report.get("status") != "passed":
        raise ValueError("新聞稿月度指標未通過驗證，停止同步")
    with Path(csv_path).open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames != COLUMNS:
            raise ValueError("新聞稿指標 CSV 欄位格式不符")
        rows = list(reader)
    if len(rows) != report.get("record_count") or [r["資料鍵"] for r in rows] != report.get("months"):
        raise ValueError("新聞稿指標 CSV 與檢核報告不一致")
    validate_metrics(rows)
    return rows


def to_sheet_row(record):
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
                    raise ValueError(f"數字無效：{col}")
                values.append(float(d))
            except InvalidOperation as exc:
                raise ValueError(f"無法轉換數值：{col}={value!r}") from exc
    return values


def existing_rows(worksheet):
    rows = worksheet.get_all_values(value_render_option="UNFORMATTED_VALUE")
    if not rows:
        return []
    header = list(rows[0])
    while header and header[-1] == "":
        header.pop()
    if header != COLUMNS:
        raise ValueError("既有『新聞稿月度指標』表頭不同，停止更新以避免覆蓋資料")
    found, keys = [], set()
    for idx, vals in enumerate(rows[1:], 2):
        if not any(str(v).strip() for v in vals):
            continue
        r = dict(zip(COLUMNS, (list(vals) + [""] * len(COLUMNS))[:len(COLUMNS)]))
        k = str(r["資料鍵"])
        if not k or k in keys:
            raise ValueError(f"既有『新聞稿月度指標』第 {idx} 列的資料鍵無效或重複")
        if k != f"{int(r['年度'])}{int(r['月份']):02d}":
            raise ValueError(f"既有『新聞稿月度指標』第 {idx} 列的資料鍵與年月不符")
        keys.add(k)
        found.append(r)
    return found


def sync(worksheet, incoming):
    validate_metrics(incoming)
    existing = existing_rows(worksheet)
    keys = {str(r["資料鍵"]) for r in incoming}
    retained = [r for r in existing if str(r["資料鍵"]) not in keys]
    merged = retained + incoming
    if len({str(r["資料鍵"]) for r in merged}) != len(merged):
        raise ValueError("新聞稿指標合併後出現重複鍵")
    merged.sort(key=lambda r: (int(r["年度"]), int(r["月份"])))
    # 先轉換完所有資料才碰到工作表。
    values = [COLUMNS] + [to_sheet_row(r) for r in merged]
    need_rows = len(values) + 20
    if worksheet.row_count < need_rows or worksheet.col_count < len(COLUMNS):
        worksheet.resize(
            rows=max(worksheet.row_count, need_rows),
            cols=max(worksheet.col_count, len(COLUMNS)),
        )
    for start in range(0, len(values), CHUNK_ROWS):
        worksheet.update(range_name=f"A{start+1}", values=values[start:start+CHUNK_ROWS], value_input_option="RAW")
    worksheet.batch_clear([f"A{len(values)+1}:{column_letter(len(COLUMNS))}{worksheet.row_count}"])
    worksheet.freeze(rows=1)
    print(f"新聞稿月度指標：本次同步 {len(incoming)} 月，保留更早歷史 {len(retained)} 月，總計 {len(merged)} 月")
    return len(merged)


def main():
    rows = read_verified_metrics()
    spreadsheet = open_spreadsheet()
    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except Exception as exc:
        from gspread.exceptions import WorksheetNotFound
        if not isinstance(exc, WorksheetNotFound):
            raise
        worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=100, cols=len(COLUMNS))
    total = sync(worksheet, rows)
    print(f"完成寫入 {WORKSHEET_NAME}：{total} 月")


if __name__ == "__main__":
    main()
