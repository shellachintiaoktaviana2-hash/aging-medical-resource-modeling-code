from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from stat_model.ahp import QUALITY_INDICATOR_MAP, run_medical_quality_ahp
from stat_model.entropy import run_entropy_evaluation
from stat_model.plotting import plot_cluster_dendrogram
from stat_model.utils import ensure_dir, read_table_auto_header


POPULATION_INDICATORS = {
    "每千人口医疗卫生机构数_计算": "positive",
    "每千人口床位数_计算": "positive",
    "每千人口卫生技术人员数_计算": "positive",
    "每千人口执业(助理)医师数_计算": "positive",
    "每千人口注册护士数_计算": "positive",
}

LAND_INDICATORS = {
    "每平方千米医疗卫生机构数": "positive",
    "每平方千米床位数": "positive",
    "每平方千米卫生技术人员数": "positive",
    "每平方千米执业(助理)医师数": "positive",
    "每平方千米注册护士数": "positive",
}


def maybe_run_land_model(path: str | None, medical_data: pd.DataFrame, output_dir: Path) -> Path | None:
    land_out = output_dir / "medical_land_entropy.xlsx"
    source_df = None
    if path:
        try:
            source_df = read_table_auto_header(path, "省级结果", ["地区"])
        except Exception:
            source_df = pd.read_excel(path)
    elif all(col in medical_data.columns for col in LAND_INDICATORS):
        source_df = medical_data

    if source_df is None:
        print("未提供土地口径原始指标，跳过土地熵权法。可用 --land-file 指定已有土地指标或结果表。")
        return None

    if all(col in source_df.columns for col in LAND_INDICATORS):
        run_entropy_evaluation(
            source_df,
            LAND_INDICATORS,
            "综合得分",
            land_out,
            group_col="四大区域",
            standardization_floor=0.1,
        )
        return land_out

    if path:
        shutil.copy2(path, land_out)
        print(f"土地口径文件未重新计算，已复制已有结果：{land_out}")
        return land_out
    return None


def maybe_run_quality_model(path: str | None, medical_data: pd.DataFrame, output_dir: Path) -> Path | None:
    quality_out = output_dir / "medical_quality_ahp.xlsx"
    if all(col in medical_data.columns for col in QUALITY_INDICATOR_MAP.values()):
        run_medical_quality_ahp(medical_data, quality_out)
        return quality_out
    if path:
        shutil.copy2(path, quality_out)
        print(f"质量口径文件未重新计算，已复制已有结果：{quality_out}")
        return quality_out
    print("未找到医疗质量AHP原始指标，跳过质量口径。可用 --quality-file 指定已有AHP结果表。")
    return None


def main() -> None:
    parser = argparse.ArgumentParser(description="Run population, land, and quality medical-resource models.")
    parser.add_argument("--medical-file", required=True, help="医疗资源核心变量 Excel")
    parser.add_argument("--sheet", default="Data_2024", help="医疗资源核心数据工作表")
    parser.add_argument("--land-file", default=None, help="可选：土地口径原始指标或已有结果 Excel")
    parser.add_argument("--quality-file", default=None, help="可选：质量口径AHP已有结果 Excel")
    parser.add_argument("--output-dir", default="outputs/medical", help="输出目录")
    args = parser.parse_args()

    output_dir = ensure_dir(args.output_dir)
    medical_data = read_table_auto_header(args.medical_file, args.sheet, ["地区"])

    population_out = output_dir / "medical_population_entropy.xlsx"
    population = run_entropy_evaluation(
        medical_data,
        POPULATION_INDICATORS,
        score_name="每千人口熵权得分",
        output_file=population_out,
    )
    plot_cluster_dendrogram(
        population["standardized"],
        list(POPULATION_INDICATORS),
        output_dir / "population_resource_cluster.png",
    )

    land_out = maybe_run_land_model(args.land_file, medical_data, output_dir)
    quality_out = maybe_run_quality_model(args.quality_file, medical_data, output_dir)

    print("医疗资源模型已完成：")
    print(f"- 人口口径熵权法：{population_out}")
    if land_out:
        print(f"- 土地口径熵权法：{land_out}")
    if quality_out:
        print(f"- 质量口径AHP：{quality_out}")


if __name__ == "__main__":
    main()
