from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from stat_model.matching import run_single_match, save_match_workbook
from stat_model.plotting import plot_gap_bar, plot_quadrant
from stat_model.utils import ensure_dir, find_column, read_table_auto_header


def read_score(path: str | Path, sheet_candidates: list[str], score_keywords: list[list[str]]) -> tuple[pd.DataFrame, str]:
    path = Path(path)
    xls = pd.ExcelFile(path)
    last_error: Exception | None = None
    for sheet in sheet_candidates + xls.sheet_names:
        if sheet not in xls.sheet_names:
            continue
        try:
            df = read_table_auto_header(path, sheet, ["地区"])
            for keys in score_keywords:
                col = find_column(df, keys, required=False)
                if col:
                    return df, col
        except Exception as exc:
            last_error = exc
    raise ValueError(f"无法从 {path} 识别得分列。最后错误：{last_error}")


def default_file(medical_dir: Path, name: str) -> Path | None:
    matches = list(medical_dir.glob(name))
    return matches[0] if matches else None


def main() -> None:
    parser = argparse.ArgumentParser(description="Run U1-U2 match analysis and create figures.")
    parser.add_argument("--aging-score", required=True, help="老龄化综合得分 Excel")
    parser.add_argument("--medical-dir", default="outputs/medical", help="医疗资源结果目录")
    parser.add_argument("--population-score", default=None, help="人口口径医疗资源得分 Excel")
    parser.add_argument("--land-score", default=None, help="土地口径医疗资源得分 Excel")
    parser.add_argument("--quality-score", default=None, help="质量口径医疗资源得分 Excel")
    parser.add_argument("--output-dir", default="outputs/matching", help="输出目录")
    args = parser.parse_args()

    output_dir = ensure_dir(args.output_dir)
    figures_dir = ensure_dir(output_dir / "figures")
    medical_dir = Path(args.medical_dir)

    aging_df, aging_col = read_score(
        args.aging_score,
        ["合成结果", "计算过程", "综合得分排名"],
        [["最终老龄化综合得分"], ["最终综合得分"], ["综合得分"], ["老龄化", "得分"]],
    )

    files = {
        "人口口径": Path(args.population_score) if args.population_score else default_file(medical_dir, "medical_population_entropy.xlsx"),
        "土地口径": Path(args.land_score) if args.land_score else default_file(medical_dir, "medical_land_entropy.xlsx"),
        "质量口径": Path(args.quality_score) if args.quality_score else default_file(medical_dir, "medical_quality_ahp.xlsx"),
    }

    score_rules = {
        "人口口径": (["综合得分排名", "每千人口熵权得分排序", "省级结果"], [["每千人口熵权得分"], ["综合得分"], ["得分"]]),
        "土地口径": (["综合得分排名", "省级结果"], [["综合得分"], ["土地", "得分"], ["得分"]]),
        "质量口径": (["省级结果", "综合得分排名"], [["综合得分"], ["AHP", "得分"], ["得分"]]),
    }

    results: dict[str, pd.DataFrame] = {}
    for name, path in files.items():
        if path is None or not path.exists():
            print(f"未找到{name}得分文件，跳过。")
            continue
        sheets, keywords = score_rules[name]
        medical_df, medical_col = read_score(path, sheets, keywords)
        result = run_single_match(aging_df, medical_df, aging_col, medical_col, name)
        results[name] = result
        out_xlsx = output_dir / f"match_{name}.xlsx"
        save_match_workbook(result, out_xlsx)
        plot_gap_bar(result, figures_dir / f"gap_{name}.png", f"{name}下老龄化-医疗资源匹配差值排序")
        plot_quadrant(result, figures_dir / f"quadrant_{name}.png", f"老龄化综合指数与{name}医疗资源指数四象限")
        print(f"{name}匹配分析完成：{out_xlsx}")

    if not results:
        raise RuntimeError("没有可用医疗资源得分文件，无法生成匹配分析。")

    summary_file = output_dir / "three_medical_views_match_summary.xlsx"
    with pd.ExcelWriter(summary_file, engine="openpyxl") as writer:
        for name, df in results.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        pd.concat(results.values(), ignore_index=True).to_excel(writer, sheet_name="三口径长表", index=False)
    print(f"三口径匹配汇总已保存：{summary_file}")


if __name__ == "__main__":
    main()
