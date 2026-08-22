from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from stat_model.ahp import AGING_DIRECTIONS, run_aging_ahp
from stat_model.entropy import run_entropy_evaluation
from stat_model.utils import ensure_dir, minmax_standardize, read_table_auto_header


def combine_aging_scores(entropy_result: pd.DataFrame, ahp_result: pd.DataFrame, output_file: Path) -> pd.DataFrame:
    entropy_score_col = "老龄化拓展模型得分"
    ahp_score_col = "AHP老龄化质量得分"
    left = entropy_result[["地区", "年份", "排名", entropy_score_col]].rename(columns={"排名": "熵权法排名", entropy_score_col: "熵权法原始得分"})
    right = ahp_result[["地区", "排名", ahp_score_col]].rename(columns={"排名": "AHP排名", ahp_score_col: "AHP原始得分"})
    df = pd.merge(left, right, on="地区", how="inner")
    df["熵权法标准化"] = minmax_standardize(df["熵权法原始得分"], fill_equal=0.0)
    df["AHP标准化"] = minmax_standardize(df["AHP原始得分"], fill_equal=0.0)
    df["最终老龄化综合得分"] = (df["熵权法标准化"] + df["AHP标准化"]) / 2
    df["最终排名"] = df["最终老龄化综合得分"].rank(ascending=False, method="min").astype(int)
    df = df.sort_values(["最终排名", "最终老龄化综合得分"], ascending=[True, False]).reset_index(drop=True)
    df.insert(0, "序号", range(1, len(df) + 1))

    params = pd.DataFrame({
        "参数": ["熵权法最小值", "熵权法最大值", "AHP最小值", "AHP最大值", "样本地区数", "熵权法合成权重", "AHP合成权重"],
        "数值": [
            df["熵权法原始得分"].min(),
            df["熵权法原始得分"].max(),
            df["AHP原始得分"].min(),
            df["AHP原始得分"].max(),
            len(df),
            0.5,
            0.5,
        ],
    })
    method = pd.DataFrame({
        "步骤": [1, 2, 3, 4],
        "处理内容": ["地区匹配", "两模型得分极差标准化", "等权平均合成", "按最终综合得分降序排名"],
        "公式/说明": ["按地区名称一一匹配", "X'=(X-min)/(max-min)", "U1=(E'+A')/2", "得分越高表示综合老龄化压力越强"],
    })
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="合成结果", index=False)
        df[["地区", "年份", "熵权法排名", "熵权法原始得分", "AHP排名", "AHP原始得分", "熵权法标准化", "AHP标准化", "最终老龄化综合得分", "最终排名"]].to_excel(
            writer,
            sheet_name="计算过程",
            index=False,
        )
        params.to_excel(writer, sheet_name="参数", index=False)
        method.to_excel(writer, sheet_name="方法说明", index=False)
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Run aging entropy, aging AHP, and composite U1 models.")
    parser.add_argument("--aging-file", required=True, help="老龄化指标 Excel，如 老龄化拓展指标数据表.xlsx")
    parser.add_argument("--sheet", default="老龄化扩展候选指标_2023", help="包含扩展老龄化指标的工作表")
    parser.add_argument("--output-dir", default="outputs/aging", help="输出目录")
    args = parser.parse_args()

    output_dir = ensure_dir(args.output_dir)
    data = read_table_auto_header(args.aging_file, args.sheet, ["地区", "年份"])

    entropy_out = output_dir / "aging_entropy_result.xlsx"
    entropy = run_entropy_evaluation(
        data,
        AGING_DIRECTIONS,
        score_name="老龄化拓展模型得分",
        output_file=entropy_out,
    )

    ahp_out = output_dir / "aging_ahp_result.xlsx"
    ahp = run_aging_ahp(data, ahp_out)

    composite_out = output_dir / "aging_composite_index.xlsx"
    composite = combine_aging_scores(entropy["result"], ahp["result"], composite_out)

    print("老龄化模型已完成：")
    print(f"- 熵权法结果：{entropy_out}")
    print(f"- AHP结果：{ahp_out}")
    print(f"- 综合U1结果：{composite_out}")
    print(composite[["地区", "最终老龄化综合得分", "最终排名"]].head(10).to_string(index=False))


if __name__ == "__main__":
    main()

