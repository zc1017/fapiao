@echo off
cd /d "%~dp0"
python invoice_app.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo 程序运行出错，请检查错误信息
    pause
)
