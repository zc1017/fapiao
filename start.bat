@echo off
cd /d "%~dp0"
python invoice_app.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo Error occurred. Press any key to exit...
    pause > nul
)
