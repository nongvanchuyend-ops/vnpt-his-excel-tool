@echo off
chcp 65001 >nul
title VNPT-HIS cURL to Excel Tool
echo ===================================================
echo     VNPT-HIS cURL to Excel Converter Tool
echo ===================================================
echo.
cd /d "%~dp0"

:: 1. Kiem tra Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [LỖI] Máy tính này chưa cài đặt Node.js!
    echo Vui lòng tải và cài đặt Node.js tại: https://nodejs.org
    echo Sau khi cài xong, bạn hãy nhấp đúp lại file run_tool.bat này.
    echo.
    pause
    exit /b
)

:: 2. Kiem tra thu muc node_modules
if not exist "node_modules\" (
    echo [THÔNG BÁO] Đang tự động cài đặt thư viện cần thiết...
    call npm install
    if %errorlevel% neq 0 (
        echo [LỖI] Không thể cài đặt thư viện. Vui lòng kiểm tra kết nối mạng.
        pause
        exit /b
    )
)

echo Đang khởi chạy máy chủ cục bộ...
echo Mở trình duyệt tại: http://localhost:3050
start "" http://localhost:3050

node server.js
pause
