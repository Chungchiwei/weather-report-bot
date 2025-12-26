@echo off
:: 切換編碼以正確顯示中文 (若需要)
chcp 65001 >nul
echo 正在設定自動執行排程...

:: 設定變數：目標檔案路徑
set "TargetFile=C:\Users\Dell\Desktop\WHL_Weather_report_system\star.bat"

:: 建立每天 08:00 的任務
:: /tn 是任務名稱
:: /tr 是要執行的檔案路徑
:: /sc daily 表示每天執行
:: /st 是時間
:: /f 表示如果任務已存在則強制覆蓋
schtasks /create /tn "WHL_Weather_Report_Morning" /tr "%TargetFile%" /sc daily /st 08:00 /f

:: 建立每天 16:00 的任務
schtasks /create /tn "WHL_Weather_Report_Afternoon" /tr "%TargetFile%" /sc daily /st 16:00 /f

echo.
echo ==========================================
echo 排程設定完成！
echo 檔案將於每天 08:00 和 16:00 自動執行。
echo ==========================================
echo 請按任意鍵退出...
pause >nul