@echo off
chcp 65001 >nul
title KBC-HP89 — App quản lý liên kết
cd /d "%~dp0"

echo ===========================================
echo   KBC-HP89 — Hệ thống quản lý liên kết
echo ===========================================
echo.

REM Tạo venv lần đầu nếu chưa có
if not exist "venv\Scripts\python.exe" (
    echo [Lần đầu] Đang tạo môi trường Python venv...
    python -m venv venv
    if errorlevel 1 (
        echo Lỗi: Không tạo được venv. Kiểm tra Python đã cài chưa.
        pause
        exit /b 1
    )
)

REM Kích hoạt venv
call "venv\Scripts\activate.bat"

REM Cài thư viện nếu chưa có flask
python -c "import flask" 2>nul
if errorlevel 1 (
    echo [Lần đầu] Đang cài thư viện cần thiết, vui lòng đợi...
    pip install --upgrade pip
    pip install -r requirements.txt
)

echo.
echo App đang chạy tại: http://127.0.0.1:5000
echo Tài khoản admin mặc định: admin / admin123
echo Bấm Ctrl+C để tắt server.
echo.

REM Mở trình duyệt sau 2 giây
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:5000"

python app.py
pause
