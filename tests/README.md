taiwan-tax-dashboard：核心三欄精簡修正版
這份修正版將原本 27 個欄位精簡為 19 欄，其中只有 3 欄是官方金額：當月實徵_億元、累計實徵_億元、全年預算_億元。其餘 16 欄為年月、統計視角、稅目階層、加總旗標、來源追溯等必要資料。月增、年增、預算差額及比率由儀表板自行計算。
必須取代 GitHub 的檔案
processor.py（改用核心三欄且接受 12 月全年表頭）
update_sheets.py（配合新版表頭寫入 19 欄，保留其他年度／月份）
tests/test_processor.py（調整三欄檢驗）
tests/test_update_sheets.py（調整欄位與覆寫檢驗）
tests/test_year_end_11412.py（新增年末回歸測試）
tests/fixtures/11412/（114 年 12 月四個 XLSX 以供測試）
crawler.py、requirements.txt 和每天早上 09:15 的排程原則上可沿用，ZIP 裡包含對應版本供檢查。
上傳順序
全部 Python 檔案與測試附件替換／新增後，才手動按 Actions → Update Taiwan Tax Revenue → Run workflow，避免新舊程式欄位不匹配。
本機測試：python -m unittest discover -s tests -v。
若 Google 試算表的「稅收原始資料」已被舊版成功寫入且有資料，程式會拒絕直接覆蓋；請先備份該工作表並進行欄位遷移，不要在未備份前直接清空。如第一次 Actions 在驗證階段失敗而尚未寫入，則通常沒有這個問題。
驗證紀錄
114 年 12 月：四張附表共 82 筆；全國當月 1701.81 億元、全年累計 37514.80 億元、全年預算 38019.39 億元。
115 年 9 月：四張附表共 82 筆；全國當月 4964.64 億元、累計 35713.84 億元、全年預算 45298.76 億元。
自動化離線測試：33 項通過。
114 年 1 月註解跨欄問題已用模擬測試；尚未用該月原始 XLSX 完整驗收。
尚未於 GitHub Actions 正式重新執行、亦尚未真的寫入 Google 試算表。
