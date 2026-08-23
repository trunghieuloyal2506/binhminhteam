@echo off
echo ================================================
echo   VinaQS Timesheet Server - Khoi dong...
echo ================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo LOI: Khong tim thay Python!
    echo Tai Python tai: https://python.org
    pause
    exit
)

REM Start server in background
start "VinaQS Timesheet Server" cmd /k "python \"%~dp0timesheet_server.py\""

REM Wait for server to start
timeout /t 2 /nobreak >nul

REM Open app in browser via local file (allows localhost calls)
echo Dang mo app...
start "" "%~dp0index.html"

echo.
echo ================================================
echo   Server dang chay tai: http://localhost:5000
echo   App da mo trong trinh duyet
echo   Nhan Ctrl+C trong cua so server de dung
echo ================================================
