"""排程健康檢查：資料是否過期，以及留下一個很少變動的狀態檔。

兩個用途：

1. check：每月 GRACE_DAY 日之後，最新月份必須已經是上個月。
   財政部通常在每月 10～15 日公布上月數；如果超過期限還沒有，常見原因是
   新聞稿標題改版而搜尋不到。這種情況其他步驟都會「成功」（只是少一個月），
   所以要靠這個檢查讓 workflow 明確失敗、寄出通知。

2. status：寫出 status/last_run.json，內容只有「最新資料月份」和「檢查的月份」，
   同一個月內不會改變，所以一個月最多產生一兩次 commit。
   公開 repo 若 60 天沒有任何活動，GitHub 會自動停用排程；這個定期 commit
   就是為了避免這件事。

使用方式：
    python health.py check                       # 過期時結束碼 1
    python health.py status --output status/last_run.json
"""

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

TAIPEI = timezone(timedelta(hours=8))

# 每月幾號之後，就必須已經有上個月的資料。
GRACE_DAY = 20


def expected_latest(today):
    """回傳 (民國年, 月)：今天起最新資料至少要到這個月份。"""

    months_back = 1 if today.day >= GRACE_DAY else 2

    year, month = today.year, today.month - months_back

    while month <= 0:
        month += 12
        year -= 1

    return year - 1911, month


def latest_period(validation):
    """從 validation.json 取出最新的 (民國年, 月)。"""

    months = validation.get("months") or []

    if not months:
        raise ValueError("validation.json 沒有任何月份")

    return max((int(m["year"]), int(m["month"])) for m in months)


def check_freshness(validation, today):
    """回傳 (是否正常, 說明文字)。"""

    latest = latest_period(validation)
    expected = expected_latest(today)

    label = f"民國 {latest[0]} 年 {latest[1]} 月"
    need = f"民國 {expected[0]} 年 {expected[1]} 月"

    if latest >= expected:
        return True, f"資料正常：最新到 {label}（今天起至少需要 {need}）"

    return False, (
        f"資料可能已過期：最新只到 {label}，"
        f"今天（{today.isoformat()}）起至少應該到 {need}。"
        "請確認財政部是否已公布，或新聞稿標題是否改版導致搜尋不到。"
    )


def build_status(validation, today):
    year, month = latest_period(validation)

    return {
        "latest_period": f"{year}{month:02d}",
        "checked_month": today.strftime("%Y-%m"),
    }


def _today():
    return datetime.now(TAIPEI).date()


def _load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv=None, today=None):
    parser = argparse.ArgumentParser(description="排程健康檢查")
    sub = parser.add_subparsers(dest="command", required=True)

    for name in ("check", "status"):
        p = sub.add_parser(name)
        p.add_argument("--validation", default="data/validation.json")

    sub.choices["status"].add_argument(
        "--output", default="status/last_run.json"
    )

    args = parser.parse_args(argv)
    today = today or _today()
    validation = _load(args.validation)

    if args.command == "check":
        ok, message = check_freshness(validation, today)
        print(message, file=sys.stdout if ok else sys.stderr)
        return 0 if ok else 1

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(build_status(validation, today), ensure_ascii=False, indent=2)
        + "\n",
        encoding="utf-8",
    )
    print(f"已寫入 {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
