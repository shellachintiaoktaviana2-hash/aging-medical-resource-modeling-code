from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from stat_model.matching import robustness_tables
from stat_model.plotting import plot_gap_heatmap
from stat_model.utils import ensure_dir


def load_match_results(match_dir: Path) -> dict[str, pd.DataFrame]:
    results = {}
    for path in sorted(match_dir.glob("match_*.xlsx")):
        name = path.stem.replace("match_", "")
        df = pd.read_excel(path, sheet_name="匹配分析结果")
        results[name] = df
    return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Run robustness checks for three medical-resource views.")
    parser.add_argument("--match-dir", default="outputs/matching", help="03_match_and_plots.py 的输出目录")
    parser.add_argument("--output-dir", default="outputs/robustness", help="输出目录")
    parser.add_argument("--base-model", default="人口口径", help="重点地区重合率的基准口径")
    args = parser.parse_args()

    output_dir = ensure_dir(args.output_dir)
    results = load_match_results(Path(args.match_dir))
    if not results:
        raise RuntimeError(f"在 {args.match_dir} 未找到 match_*.xlsx。")

    tables = robustness_tables(results, base_model=args.base_model)
    output_file = output_dir / "robustness_checks.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet, table in tables.items():
            table.to_excel(writer, sheet_name=sheet[:31], index=True if "矩阵" in sheet else False)

    compare = tables["三口径对比总表"]
    plot_gap_heatmap(compare, output_dir / "three_views_gap_heatmap.png")
    print(f"稳健性检验已完成：{output_file}")


if __name__ == "__main__":
    main()
