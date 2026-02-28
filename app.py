import argparse

from analyzer import run_analysis


def main():
    parser = argparse.ArgumentParser(description="标准化财务报表自动分析")
    parser.add_argument("--input", required=True, help="输入Excel文件路径")
    parser.add_argument("--output-dir", required=True, help="输出目录")
    parser.add_argument("--sheet-bs", default="资产负债表", help="资产负债表sheet名")
    parser.add_argument("--sheet-is", default="损益表 ", help="损益表sheet名（默认带尾随空格）")
    parser.add_argument("--sheet-cf", default="现金流量表", help="现金流量表sheet名")
    parser.add_argument("--sheet-kpi", default="主要财务数据及指标", help="可选指标校验sheet名")
    parser.add_argument("--config", default="config.yml", help="配置文件路径")
    args = parser.parse_args()

    run_analysis(
        input_file=args.input,
        output_dir=args.output_dir,
        sheet_bs=args.sheet_bs,
        sheet_is=args.sheet_is,
        sheet_cf=args.sheet_cf,
        sheet_kpi=args.sheet_kpi,
        config_path=args.config,
    )
    print(f"分析完成，结果已输出到：{args.output_dir}")


if __name__ == "__main__":
    main()
