from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd

from .utils import ensure_dir, minmax_standardize, normalize_province_column


def entropy_weight(std_df: pd.DataFrame) -> tuple[pd.Series, pd.Series, pd.Series, pd.DataFrame]:
    """Calculate entropy weights from a standardized positive-indicator matrix."""
    X = std_df.astype(float).clip(lower=0)
    n = len(X)
    if n <= 1:
        raise ValueError("熵权法至少需要两个评价对象。")
    P = X / (X.sum(axis=0) + 1e-12)
    k = 1 / math.log(n)
    with np.errstate(divide="ignore", invalid="ignore"):
        entropy = -k * np.where(P > 0, P * np.log(P), 0.0).sum(axis=0)
    entropy = pd.Series(entropy, index=X.columns, name="熵值")
    diff = 1 - entropy
    if np.isclose(float(diff.sum()), 0):
        weight = pd.Series(1 / len(diff), index=X.columns, name="权重")
    else:
        weight = diff / diff.sum()
        weight.name = "权重"
    diff.name = "差异系数"
    return weight, entropy, diff, P


def run_entropy_evaluation(
    data: pd.DataFrame,
    indicators: dict[str, str],
    score_name: str,
    output_file: str | Path,
    group_col: str | None = None,
    standardization_floor: float = 0.0,
) -> dict[str, pd.DataFrame]:
    df = normalize_province_column(data, "地区")
    needed = ["地区"] + list(indicators.keys())
    if group_col and group_col in df.columns:
        needed.insert(1, group_col)
    if "年份" in df.columns and "年份" not in needed:
        needed.insert(1, "年份")
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"熵权法输入缺少列：{missing}")

    raw = df[needed].copy()
    std = raw[["地区"] + ([group_col] if group_col and group_col in raw.columns else [])].copy()
    if "年份" in raw.columns and "年份" not in std.columns:
        std["年份"] = raw["年份"]
    for col, direction in indicators.items():
        z = minmax_standardize(raw[col], direction=direction, fill_equal=1.0)
        if standardization_floor:
            z = standardization_floor + (1 - standardization_floor) * z
        std[col] = z

    matrix = std[list(indicators.keys())]
    weight, entropy, diff, proportion = entropy_weight(matrix)
    score = matrix.dot(weight)

    result_cols = ["地区"]
    if "年份" in raw.columns:
        result_cols.append("年份")
    if group_col and group_col in raw.columns:
        result_cols.append(group_col)

    result = raw[result_cols].copy()
    result[score_name] = score
    result["排名"] = result[score_name].rank(ascending=False, method="min").astype(int)
    result = result.sort_values(["排名", score_name], ascending=[True, False]).reset_index(drop=True)

    weight_table = pd.DataFrame({
        "指标": list(indicators.keys()),
        "指标方向": [indicators[c] for c in indicators],
        "标准化下限": standardization_floor,
        "熵值": entropy.values,
        "差异系数": diff.values,
        "权重": weight.values,
    })

    proportion_out = raw[["地区"]].join(proportion.reset_index(drop=True))
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="原始数据", index=False)
        std.to_excel(writer, sheet_name="标准化矩阵", index=False)
        proportion_out.to_excel(writer, sheet_name="比重矩阵P", index=False)
        weight_table.to_excel(writer, sheet_name="指标权重", index=False)
        result.to_excel(writer, sheet_name="综合得分排名", index=False)

    return {
        "raw": raw,
        "standardized": std,
        "weights": weight_table,
        "result": result,
    }
