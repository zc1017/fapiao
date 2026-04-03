@echo off
cd /d "%~dp0"
set FLAGS_use_mkldnn=0
set FLAGS_enable_onednn_backend=0
set MKL_DEBUG_CPU_TYPE=0
set OMP_NUM_THREADS=1
python invoice_app.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo 程序运行出错，请检查错误信息
    pause
)
