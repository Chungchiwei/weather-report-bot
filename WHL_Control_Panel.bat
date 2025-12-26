@echo off
chcp 65001 >nul
title WHL 天氣報告排程控制台
color 0B

:MENU
cls
echo ========================================================
echo           WHL 天氣報告系統 - 排程控制台
echo ========================================================
echo.
echo    目標檔案: C:\Users\Dell\Desktop\WHL_Weather_report_system\star.bat
echo.
echo    [1] 啟用標準排程 (每天 08:00 和 16:00 自動執行)
echo    [2] 自訂單一時間 (輸入指定時間，每天執行)
echo    [3] 暫停/取消發送 (移除所有排程)
echo    [4] 檢查目前排程狀態
echo    [Q] 離開
echo.
echo ========================================================
set /p choice="請選擇功能代號 [1-4, Q]: "

if "%choice%"=="1" goto SET_STANDARD
if "%choice%"=="2" goto SET_CUSTOM
if "%choice%"=="3" goto STOP_TASK
if "%choice%"=="4" goto CHECK_STATUS
if /i "%choice%"=="Q" exit
goto MENU

:: ========================================================
:: 1. 設定標準排程 (08:00 & 16:00)
:: ========================================================
:SET_STANDARD
cls
echo 正在設定標準排程 (08:00 及 16:00)...
set "TargetFile=C:\Users\Dell\Desktop\WHL_Weather_report_system\star.bat"

:: 先刪除舊的以防重複
schtasks /delete /tn "WHL_Weather_Morning" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Afternoon" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Custom" /f >nul 2>&1

:: 建立新任務
schtasks /create /tn "WHL_Weather_Morning" /tr "%TargetFile%" /sc daily /st 08:00 /f
if %errorlevel% neq 0 goto ERROR_MSG

schtasks /create /tn "WHL_Weather_Afternoon" /tr "%TargetFile%" /sc daily /st 16:00 /f
if %errorlevel% neq 0 goto ERROR_MSG

echo.
echo [成功] 已設定每天 08:00 與 16:00 自動發送！
pause
goto MENU

:: ========================================================
:: 2. 自訂時間
:: ========================================================
:SET_CUSTOM
cls
echo 請輸入想要每天執行的時間 (格式 HH:MM，例如 09:30 或 14:00)
set /p UserTime="請輸入時間: "
set "TargetFile=C:\Users\Dell\Desktop\WHL_Weather_report_system\star.bat"

:: 刪除標準排程以免衝突 (可根據需求保留，這裡設定為取代)
schtasks /delete /tn "WHL_Weather_Morning" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Afternoon" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Custom" /f >nul 2>&1

schtasks /create /tn "WHL_Weather_Custom" /tr "%TargetFile%" /sc daily /st %UserTime% /f
if %errorlevel% neq 0 (
    echo.
    echo [失敗] 時間格式錯誤或權限不足，請確保格式為 HH:MM。
    pause
    goto MENU
)

echo.
echo [成功] 已設定每天 %UserTime% 自動發送！
pause
goto MENU

:: ========================================================
:: 3. 暫停/移除排程
:: ========================================================
:STOP_TASK
cls
echo 正在移除自動發送排程...

schtasks /delete /tn "WHL_Weather_Morning" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Afternoon" /f >nul 2>&1
schtasks /delete /tn "WHL_Weather_Custom" /f >nul 2>&1

echo.
echo [已暫停] 所有自動發送任務已移除，系統將不再自動執行。
pause
goto MENU

:: ========================================================
:: 4. 檢查狀態
:: ========================================================
:CHECK_STATUS
cls
echo 目前系統中的 WHL 相關排程：
echo --------------------------------------------------------
schtasks /query | findstr "WHL_Weather"
if %errorlevel% neq 0 echo 目前沒有設定任何 WHL 自動排程。
echo --------------------------------------------------------
echo.
pause
goto MENU

:ERROR_MSG
echo.
echo [錯誤] 設定失敗！
echo 請確認您是否已按右鍵選擇「以系統管理員身分執行」。
pause
goto MENU