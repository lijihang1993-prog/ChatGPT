# 财务报表自动分析程序

## 功能
- 读取标准化 Excel 财务报表（资产负债表、损益表、现金流量表）。
- 自动做科目名称标准化匹配（去空格、全角空格、冒号及常见符号）。
- 计算规模增长、盈利、偿债、现金质量、营运相关指标。
- 生成中文 Markdown 分析报告（含“口径缺失/数据不可得”提示、风险规则触发）。
- 导出指标表 `metrics.xlsx` 与趋势图（收入/净利润）。

## 安装
```bash
pip install -r requirements.txt
```

## 运行
```bash
python app.py --input 财务报表.xlsx --output-dir output/
```

可选参数：
- `--sheet-bs`（默认：`资产负债表`）
- `--sheet-is`（默认：`损益表 `，末尾空格）
- `--sheet-cf`（默认：`现金流量表`）
- `--config`（默认：`config.yml`）

## 输出
- `output/report.md`
- `output/metrics.xlsx`
- `output/charts/revenue_trend.png`
- `output/charts/net_profit_trend.png`

## 测试
```bash
pytest -q
```
