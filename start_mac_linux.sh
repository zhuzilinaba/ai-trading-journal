#!/bin/bash
echo "检查依赖..."
pip3 install openpyxl pandas pillow -q 2>/dev/null || pip install openpyxl pandas pillow -q

echo "启动交易复盘日志..."
python3 "$(dirname "$0")/trade_journal.py"
