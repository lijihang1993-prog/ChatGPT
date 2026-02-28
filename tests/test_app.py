from pathlib import Path

import pandas as pd

from analyzer import (
    compute_metrics,
    normalize_subject_name,
    read_statement_sheet,
    run_analysis,
    validate_with_optional_kpi,
)


def create_sample_excel(path: Path, with_kpi: bool = True):
    bs = pd.DataFrame(
        {
            "项目": ["资 产 总 计", "负债合计", "流动资产合计", "流动负债合计", "应收账款", "存货"],
            "2023年": [1000, 400, 500, 250, 100, 80],
            "2024年": [1100, 500, 560, 300, 130, 120],
            "2025年": [1200, 620, 600, 350, 170, 140],
        }
    )
    ins = pd.DataFrame(
        {
            "项目": ["营业收入", "营业成本", "净利润", "销售费用", "管理费用", "研发费用", "财务费用"],
            "2023年": [800, 500, 80, 30, 20, 10, 5],
            "2024年": [900, 580, 90, 32, 21, 12, 5],
            "2025年": [850, 570, 70, 35, 23, 12, 6],
        }
    )
    cf = pd.DataFrame(
        {
            "项目": ["经营活动产生的现金流量净额"],
            "2023年": [70],
            "2024年": [50],
            "2025年": [20],
        }
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        bs.to_excel(writer, sheet_name="资产负债表", index=False)
        ins.to_excel(writer, sheet_name="损益表 ", index=False)
        cf.to_excel(writer, sheet_name="现金流量表", index=False)
        if with_kpi:
            kpi = pd.DataFrame(
                {
                    "项目": ["营业收入", "净利润"],
                    "2023年": [0.08, 0.008],
                    "2024年": [0.09, 0.009],
                    "2025年": [0.085, 0.007],
                }
            )
            kpi.to_excel(writer, sheet_name="主要财务数据及指标", index=False)


CONFIG = {
    "mappings": {
        "revenue": ["营业收入"],
        "operating_cost": ["营业成本"],
        "net_profit": ["净利润"],
        "selling_expense": ["销售费用"],
        "admin_expense": ["管理费用"],
        "rd_expense": ["研发费用"],
        "finance_expense": ["财务费用"],
        "total_assets": ["资产总计"],
        "total_liabilities": ["负债合计"],
        "current_assets": ["流动资产合计"],
        "current_liabilities": ["流动负债合计"],
        "operating_cashflow": ["经营活动产生的现金流量净额"],
        "accounts_receivable": ["应收账款"],
        "inventory": ["存货"],
    }
}


def test_normalize_subject_name():
    assert normalize_subject_name("资 产：总 计") == "资产总计"


def test_read_and_compute_metrics(tmp_path):
    xlsx = tmp_path / "财务报表.xlsx"
    create_sample_excel(xlsx)

    bs_df = read_statement_sheet(str(xlsx), "资产负债表")
    is_df = read_statement_sheet(str(xlsx), "损益表 ")
    cf_df = read_statement_sheet(str(xlsx), "现金流量表")

    metrics = compute_metrics(bs_df, is_df, cf_df, CONFIG)
    assert len(metrics) == 3
    assert metrics.loc[2, "营业收入"] == 850
    assert round(metrics.loc[2, "资产负债率"], 4) == round(620 / 1200, 4)


def test_optional_kpi_validation_detects_scale(tmp_path):
    xlsx = tmp_path / "财务报表.xlsx"
    create_sample_excel(xlsx)

    bs_df = read_statement_sheet(str(xlsx), "资产负债表")
    is_df = read_statement_sheet(str(xlsx), "损益表 ")
    cf_df = read_statement_sheet(str(xlsx), "现金流量表")
    kpi_df = read_statement_sheet(str(xlsx), "主要财务数据及指标")

    metrics = compute_metrics(bs_df, is_df, cf_df, CONFIG)
    notes = validate_with_optional_kpi(metrics, kpi_df, CONFIG)
    assert any("10000" in n for n in notes)


def test_run_analysis_outputs(tmp_path):
    xlsx = tmp_path / "财务报表.xlsx"
    create_sample_excel(xlsx)

    config_path = tmp_path / "config.yml"
    config_path.write_text(Path("config.yml").read_text(encoding="utf-8"), encoding="utf-8")

    out_dir = tmp_path / "output"
    run_analysis(
        input_file=str(xlsx),
        output_dir=str(out_dir),
        sheet_bs="资产负债表",
        sheet_is="损益表 ",
        sheet_cf="现金流量表",
        sheet_kpi="主要财务数据及指标",
        config_path=str(config_path),
    )

    assert (out_dir / "report.md").exists()
    assert (out_dir / "metrics.xlsx").exists()
    assert (out_dir / "charts" / "revenue_trend.png").exists()
    assert (out_dir / "charts" / "net_profit_trend.png").exists()
