@echo off
chcp 65001 >nul
echo 正在启动交易复盘日志系统...
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b
)

:: 安装依赖
echo 检查依赖包...
pip install openpyxl pandas pillow -q

:: 启动程序
python trade_journal.py

pause
