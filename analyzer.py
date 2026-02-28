from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import pandas as pd
import yaml

PUNCT_PATTERN = re.compile(r"[\s\u3000:：()（）\-—_、,，.。/\\]")


@dataclass
class AnalysisContext:
    periods: List[str]
    metrics_df: pd.DataFrame
    report_text: str
    validation_notes: List[str]


def load_config(config_path: str = "config.yml") -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def normalize_subject_name(text: str) -> str:
    if text is None:
        return ""
    return PUNCT_PATTERN.sub("", str(text)).strip().lower()


def _parse_numeric(value) -> Optional[float]:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if s in {"", "-", "--", "nan", "None"}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _normalize_columns(columns: List[str]) -> List[str]:
    return [str(c).strip() for c in columns]


def read_statement_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = _normalize_columns(list(df.columns))
    if df.shape[1] < 2:
        raise ValueError(f"工作表 {sheet_name} 列数不足，至少应包含项目列与期间列")

    first_col = df.columns[0]
    period_cols = list(df.columns[1:])
    rows = []
    for _, row in df.iterrows():
        item_raw = row[first_col]
        if pd.isna(item_raw):
            continue
        item = str(item_raw).strip()
        norm_item = normalize_subject_name(item)
        values = {str(col): _parse_numeric(row[col]) for col in period_cols}
        rows.append({"item": item, "norm_item": norm_item, **values})

    return pd.DataFrame(rows)


def try_read_optional_sheet(file_path: str, sheet_name: str) -> Optional[pd.DataFrame]:
    xls = pd.ExcelFile(file_path)
    wanted_norm = normalize_subject_name(sheet_name)
    for actual in xls.sheet_names:
        if normalize_subject_name(actual) == wanted_norm:
            return read_statement_sheet(file_path, actual)
    return None


def _find_item_values(df: pd.DataFrame, aliases: List[str], periods: List[str]) -> Dict[str, Optional[float]]:
    norm_aliases = {normalize_subject_name(a) for a in aliases}
    matched = df[df["norm_item"].isin(norm_aliases)]
    if matched.empty:
        return {p: None for p in periods}
    row = matched.iloc[0]
    return {p: _parse_numeric(row.get(p)) for p in periods}


def _pct_change(curr: Optional[float], prev: Optional[float]) -> Optional[float]:
    if curr is None or prev is None or prev == 0:
        return None
    return (curr - prev) / abs(prev)


def _safe_div(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b in (None, 0):
        return None
    return a / b


def compute_metrics(bs_df: pd.DataFrame, is_df: pd.DataFrame, cf_df: pd.DataFrame, config: dict) -> pd.DataFrame:
    periods = [c for c in is_df.columns if c not in {"item", "norm_item"}]
    mappings = config["mappings"]

    rev = _find_item_values(is_df, mappings["revenue"], periods)
    cost = _find_item_values(is_df, mappings["operating_cost"], periods)
    npf = _find_item_values(is_df, mappings["net_profit"], periods)
    sell = _find_item_values(is_df, mappings["selling_expense"], periods)
    adm = _find_item_values(is_df, mappings["admin_expense"], periods)
    rd = _find_item_values(is_df, mappings["rd_expense"], periods)
    fin = _find_item_values(is_df, mappings["finance_expense"], periods)

    ta = _find_item_values(bs_df, mappings["total_assets"], periods)
    tl = _find_item_values(bs_df, mappings["total_liabilities"], periods)
    ca = _find_item_values(bs_df, mappings["current_assets"], periods)
    cl = _find_item_values(bs_df, mappings["current_liabilities"], periods)
    ar = _find_item_values(bs_df, mappings["accounts_receivable"], periods)
    inv = _find_item_values(bs_df, mappings["inventory"], periods)

    ocf = _find_item_values(cf_df, mappings["operating_cashflow"], periods)

    records = []
    for i, p in enumerate(periods):
        prev = periods[i - 1] if i > 0 else None
        rev_growth = _pct_change(rev[p], rev[prev]) if prev else None
        np_growth = _pct_change(npf[p], npf[prev]) if prev else None
        ar_growth = _pct_change(ar[p], ar[prev]) if prev else None
        inv_growth = _pct_change(inv[p], inv[prev]) if prev else None

        expense_sum = sum(x for x in [sell[p], adm[p], rd[p], fin[p]] if x is not None)
        expense_available = any(x is not None for x in [sell[p], adm[p], rd[p], fin[p]])

        records.append(
            {
                "期间": p,
                "营业收入": rev[p],
                "营业收入同比": rev_growth,
                "净利润": npf[p],
                "净利润同比": np_growth,
                "营业成本": cost[p],
                "毛利率": _safe_div((rev[p] - cost[p]) if rev[p] is not None and cost[p] is not None else None, rev[p]),
                "期间费用率": _safe_div(expense_sum if expense_available else None, rev[p]),
                "净利率": _safe_div(npf[p], rev[p]),
                "资产总计": ta[p],
                "负债合计": tl[p],
                "流动资产合计": ca[p],
                "流动负债合计": cl[p],
                "资产负债率": _safe_div(tl[p], ta[p]),
                "流动比率": _safe_div(ca[p], cl[p]),
                "经营活动现金流量净额": ocf[p],
                "经营现金流/净利润": _safe_div(ocf[p], npf[p]),
                "应收账款": ar[p],
                "应收账款同比": ar_growth,
                "存货": inv[p],
                "存货同比": inv_growth,
            }
        )

    return pd.DataFrame(records)


def _fmt_num(v: Optional[float], unit_divisor: float) -> str:
    if v is None or pd.isna(v):
        return "口径缺失/数据不可得"
    return f"{v / unit_divisor:,.2f}"


def _fmt_pct(v: Optional[float], digits: int = 1) -> str:
    if v is None or pd.isna(v):
        return "口径缺失/数据不可得"
    return f"{v * 100:.{digits}f}%"


def _fmt_ratio(v: Optional[float]) -> str:
    if v is None or pd.isna(v):
        return "口径缺失/数据不可得"
    return f"{v:.2f}"


def _unit_divisor(unit: str) -> float:
    return {"元": 1.0, "万元": 1e4, "百万": 1e6}.get(unit, 1.0)


def _estimate_scale_factor(metrics_df: pd.DataFrame, kpi_df: Optional[pd.DataFrame], config: dict) -> float:
    if kpi_df is None:
        return 1.0
    periods = metrics_df["期间"].astype(str).tolist()
    if not periods:
        return 1.0
    kpi_revenue = _find_item_values(kpi_df, config["mappings"]["revenue"], periods)
    base_period = periods[-1]
    report_revenue = metrics_df.iloc[-1]["营业收入"]
    kpi_rev = kpi_revenue.get(base_period)
    if report_revenue in (None, 0) or kpi_rev in (None, 0):
        return 1.0
    ratio = report_revenue / kpi_rev
    if abs(ratio - 10000) < 5000:
        return 10000.0
    if abs(ratio - 0.0001) < 0.00005:
        return 0.0001
    return 1.0


def validate_with_optional_kpi(metrics_df: pd.DataFrame, kpi_df: Optional[pd.DataFrame], config: dict) -> List[str]:
    if kpi_df is None:
        return ["未提供“主要财务数据及指标”工作表，未执行交叉校验。"]

    periods = metrics_df["期间"].astype(str).tolist()
    scale = _estimate_scale_factor(metrics_df, kpi_df, config)
    notes: List[str] = []
    if scale != 1.0:
        notes.append(f"检测到主表与指标表单位可能存在比例差，校验时按 {scale:g} 倍进行换算。")

    revenue_kpi = _find_item_values(kpi_df, config["mappings"]["revenue"], periods)
    net_profit_kpi = _find_item_values(kpi_df, config["mappings"]["net_profit"], periods)

    for col, kpi_values, label in [
        ("营业收入", revenue_kpi, "营业收入"),
        ("净利润", net_profit_kpi, "净利润"),
    ]:
        for _, row in metrics_df.iterrows():
            p = row["期间"]
            left = row[col]
            right = kpi_values.get(p)
            if left is None or right is None:
                continue
            right_scaled = right * scale
            if left == 0:
                continue
            diff_ratio = abs(left - right_scaled) / abs(left)
            if diff_ratio > 0.05:
                notes.append(
                    f"{p}{label}在主表与指标表存在偏差：主表{left:.2f}，指标表换算后{right_scaled:.2f}，差异{diff_ratio * 100:.1f}%。"
                )

    if not notes:
        notes.append("主要财务数据及指标交叉校验未发现明显差异。")
    return notes


def detect_risks(metrics_df: pd.DataFrame, config: dict) -> List[str]:
    thresholds = config.get("thresholds", {})
    risks = []
    if len(metrics_df) < 2:
        return risks

    latest = metrics_df.iloc[-1]
    prev = metrics_df.iloc[-2]

    rev_g = latest.get("营业收入同比")
    gm = latest.get("毛利率")
    gm_prev = prev.get("毛利率")
    if pd.notna(rev_g) and rev_g < 0 and pd.notna(gm) and pd.notna(gm_prev) and gm < gm_prev:
        risks.append(
            f"收入下降但成本下降不足导致毛利率恶化：{latest['期间']}收入同比{_fmt_pct(rev_g)}，毛利率由{_fmt_pct(gm_prev)}降至{_fmt_pct(gm)}。"
        )

    ocf_np = latest.get("经营现金流/净利润")
    min_ratio = thresholds.get("ocf_to_net_profit_min", 0.6)
    ocf_value = latest.get("经营活动现金流量净额")
    if pd.notna(ocf_np) and (ocf_np < min_ratio or (pd.notna(ocf_value) and ocf_value < 0)):
        risks.append(f"经营现金流显著低于净利润或为负：{latest['期间']}经营现金流/净利润{_fmt_pct(ocf_np)}，阈值{min_ratio:.2f}。")

    debt = latest.get("资产负债率")
    debt_prev = prev.get("资产负债率")
    debt_increase_threshold = thresholds.get("debt_ratio_increase_pct_point", 5.0) / 100
    if pd.notna(debt) and pd.notna(debt_prev) and (debt - debt_prev) > debt_increase_threshold:
        risks.append(f"资产负债率上升明显：由{_fmt_pct(debt_prev)}升至{_fmt_pct(debt)}，上升{(debt - debt_prev) * 100:.1f}pct。")

    gap_threshold = thresholds.get("ar_or_inventory_vs_revenue_growth_gap_pct_point", 10.0) / 100
    for field in ["应收账款同比", "存货同比"]:
        fg = latest.get(field)
        if pd.notna(fg) and pd.notna(rev_g) and (fg - rev_g) > gap_threshold:
            name = field.replace("同比", "")
            risks.append(f"{name}增长显著高于收入：{latest['期间']}{name}同比{_fmt_pct(fg)}，收入同比{_fmt_pct(rev_g)}。")

    return risks


def generate_report(metrics_df: pd.DataFrame, config: dict, validation_notes: List[str]) -> str:
    unit = config.get("display", {}).get("unit", "百万")
    divisor = _unit_divisor(unit)
    latest = metrics_df.iloc[-1]

    lines: List[str] = ["# 自动财务分析报告", "", "## 摘要"]
    lines.extend(
        [
            f"- {latest['期间']}营业收入{_fmt_num(latest['营业收入'], divisor)}{unit}，同比{_fmt_pct(latest['营业收入同比'])}；净利润{_fmt_num(latest['净利润'], divisor)}{unit}，同比{_fmt_pct(latest['净利润同比'])}。",
            f"- {latest['期间']}毛利率{_fmt_pct(latest['毛利率'])}，期间费用率{_fmt_pct(latest['期间费用率'])}，净利率{_fmt_pct(latest['净利率'])}。",
            f"- {latest['期间']}资产负债率{_fmt_pct(latest['资产负债率'])}，流动比率{_fmt_ratio(latest['流动比率'])}。",
            f"- {latest['期间']}经营活动现金流量净额{_fmt_num(latest['经营活动现金流量净额'], divisor)}{unit}，经营现金流/净利润{_fmt_pct(latest['经营现金流/净利润'])}。",
        ]
    )

    lines.extend(["", "## 经营表现（收入/成本/利润与驱动）"])
    for _, row in metrics_df.iterrows():
        lines.append(
            f"- {row['期间']}：收入{_fmt_num(row['营业收入'], divisor)}{unit}（同比{_fmt_pct(row['营业收入同比'])}），营业成本{_fmt_num(row['营业成本'], divisor)}{unit}，净利润{_fmt_num(row['净利润'], divisor)}{unit}（同比{_fmt_pct(row['净利润同比'])}），毛利率{_fmt_pct(row['毛利率'])}。"
        )

    lines.extend(["", "## 财务结构（资产、负债、权益变动与结构性变化）"])
    for _, row in metrics_df.iterrows():
        equity = None
        if pd.notna(row["资产总计"]) and pd.notna(row["负债合计"]):
            equity = row["资产总计"] - row["负债合计"]
        lines.append(
            f"- {row['期间']}：资产总计{_fmt_num(row['资产总计'], divisor)}{unit}，负债合计{_fmt_num(row['负债合计'], divisor)}{unit}，权益估算{_fmt_num(equity, divisor)}{unit}，资产负债率{_fmt_pct(row['资产负债率'])}，流动比率{_fmt_ratio(row['流动比率'])}。"
        )

    lines.extend(["", "## 现金流与质量（经营现金流、含金量）"])
    for _, row in metrics_df.iterrows():
        lines.append(
            f"- {row['期间']}：经营活动现金流量净额{_fmt_num(row['经营活动现金流量净额'], divisor)}{unit}，经营现金流/净利润{_fmt_pct(row['经营现金流/净利润'])}。"
        )

    lines.extend(["", "## 风险与关注点（规则触发）"])
    risks = detect_risks(metrics_df, config)
    if risks:
        lines.extend([f"- {x}" for x in risks])
    else:
        lines.append("- 未触发预设风险规则。")

    lines.extend(["", "## 交叉校验（可选：主要财务数据及指标）"])
    lines.extend([f"- {x}" for x in validation_notes])

    lines.extend(["", "## 附录：指标表", "", metrics_df.to_markdown(index=False)])
    return "\n".join(lines)


def save_charts(metrics_df: pd.DataFrame, output_dir: str) -> Tuple[str, str]:
    chart_dir = os.path.join(output_dir, "charts")
    os.makedirs(chart_dir, exist_ok=True)

    periods = metrics_df["期间"].astype(str).tolist()
    rev_path = os.path.join(chart_dir, "revenue_trend.png")
    np_path = os.path.join(chart_dir, "net_profit_trend.png")

    def _plot(y_col: str, title: str, path: str):
        plt.figure(figsize=(8, 4))
        plt.plot(periods, metrics_df[y_col], marker="o")
        plt.title(title)
        plt.xlabel("期间")
        plt.ylabel(y_col)
        plt.grid(alpha=0.3)
        plt.tight_layout()
        plt.savefig(path, dpi=150)
        plt.close()

    _plot("营业收入", "营业收入趋势", rev_path)
    _plot("净利润", "净利润趋势", np_path)
    return rev_path, np_path


def run_analysis(
    input_file: str,
    output_dir: str,
    sheet_bs: str,
    sheet_is: str,
    sheet_cf: str,
    config_path: str = "config.yml",
    sheet_kpi: str = "主要财务数据及指标",
) -> AnalysisContext:
    config = load_config(config_path)
    bs_df = read_statement_sheet(input_file, sheet_bs)
    is_df = read_statement_sheet(input_file, sheet_is)
    cf_df = read_statement_sheet(input_file, sheet_cf)
    kpi_df = try_read_optional_sheet(input_file, sheet_kpi)

    metrics_df = compute_metrics(bs_df, is_df, cf_df, config)
    validation_notes = validate_with_optional_kpi(metrics_df, kpi_df, config)
    report = generate_report(metrics_df, config, validation_notes)

    os.makedirs(output_dir, exist_ok=True)
    metrics_df.to_excel(os.path.join(output_dir, "metrics.xlsx"), index=False)
    with open(os.path.join(output_dir, "report.md"), "w", encoding="utf-8") as f:
        f.write(report)
    save_charts(metrics_df, output_dir)

    return AnalysisContext(
        periods=metrics_df["期间"].astype(str).tolist(),
        metrics_df=metrics_df,
        report_text=report,
        validation_notes=validation_notes,
    )
