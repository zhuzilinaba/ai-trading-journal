# 交易复盘日志系统

一个本地桌面程序，帮助记录、复盘、导出交易数据。

## 功能

- ✅ 手动录入交易信息（品种、方向、价格、数量、盈亏等）
- ✅ 上传交割单截图 → Claude AI 自动识别交易信息
- ✅ 复盘笔记（入场理由、离场理由、总结、教训）
- ✅ 情绪状态 + 交易评级（A/B/C/D/F）
- ✅ 导出 Excel（含统计分析、每日盈亏、累计曲线图）
- ✅ 搜索过滤历史记录
- ✅ 数据本地存储（JSON文件，在 ~/.trade_journal_data.json）

## 快速启动

### Windows
```
双击 start_windows.bat
```

### Mac / Linux
```bash
python3 trade_journal.py
```

## 依赖安装

```bash
pip install openpyxl pandas pillow
```

## 图片识别功能（可选）

需要 Claude API Key：
1. 访问 https://console.anthropic.com/ 注册获取
2. 程序内点击 "⚙ API设置" 填入
3. 之后上传交割单图片点"🔍 AI识别"即可自动填入字段

## Excel 导出格式

导出的 Excel 包含三个 Sheet：
- **交易记录**：所有交易详情，颜色区分买入/卖出/盈亏
- **统计分析**：胜率、盈亏比、情绪分布、评级分布（自动计算）
- **每日盈亏**：按日汇总 + 累计盈亏曲线图

## 数据说明

所有数据保存在本地 `~/.trade_journal_data.json`，不会上传到任何服务器。
API Key 单独保存在 `~/.trade_journal_api_key.txt`。
