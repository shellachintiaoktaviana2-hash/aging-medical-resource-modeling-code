from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .utils import ensure_dir, minmax_standardize, normalize_province_column


def classify_match(diff: float) -> str:
    if diff < -0.20:
        return "严重失配"
    if diff < -0.05:
        return "相对失配"
    if diff <= 0.05:
        return "基本匹配"
    if diff <= 0.20:
        return "相对超前"
    return "明显超前"


def quadrant_type(aging: float, medical: float, aging_boundary: float, medical_boundary: float) -> str:
    if aging >= aging_boundary and medical >= medical_boundary:
        return "高老龄-高资源"
    if aging >= aging_boundary and medical < medical_boundary:
        return "高老龄-低资源"
    if aging < aging_boundary and medical < medical_boundary:
        return "低老龄-低资源"
    return "低老龄-高资源"


def coupling_degree(u1: float, u2: float) -> float:
    if u1 < 0 or u2 < 0 or np.isclose(u1 + u2, 0):
        return np.nan
    return float(2 * np.sqrt(u1 * u2) / (u1 + u2))


def priority_label(row: pd.Series) -> str:
    if row["匹配差值"] < -0.05 and row["四象限类型"] == "高老龄-低资源":
        return "第一优先级"
    if row["匹配差值"] < -0.05:
        return "第二优先级"
    if row["四象限类型"] == "低老龄-低资源":
        return "第三优先级"
    return "一般关注"


def prepare_index(df: pd.DataFrame, score_col: str, output_col: str) -> pd.DataFrame:
    out = normalize_province_column(df, "地区")
    if score_col not in out.columns:
        raise ValueError(f"找不到得分列 {score_col}，当前列名：{list(out.columns)}")
    out = out[["地区", score_col]].copy()
    out[output_col] = pd.to_numeric(out[score_col], errors="coerce")
    if out[output_col].min() < 0 or out[output_col].max() > 1:
        out[f"{output_col}_原始"] = out[output_col]
        out[output_col] = minmax_standardize(out[output_col], fill_equal=0.0)
    return out.drop(columns=[score_col]).dropna(subset=[output_col])


def run_single_match(
    aging_df: pd.DataFrame,
    medical_df: pd.DataFrame,
    aging_score_col: str,
    medical_score_col: str,
    model_name: str,
) -> pd.DataFrame:
    u1 = prepare_index(aging_df, aging_score_col, "老龄化指数")
    u2 = prepare_index(medical_df, medical_score_col, "医疗资源指数")
    df = pd.merge(u1, u2, on="地区", how="inner")
    if df.empty:
        raise ValueError(f"{model_name} 未匹配到共同地区。")

    df["匹配差值"] = df["医疗资源指数"] - df["老龄化指数"]
    df["匹配类型"] = df["匹配差值"].apply(classify_match)
    aging_boundary = df["老龄化指数"].mean()
    medical_boundary = df["医疗资源指数"].mean()
    df["四象限类型"] = df.apply(lambda r: quadrant_type(r["老龄化指数"], r["医疗资源指数"], aging_boundary, medical_boundary), axis=1)
    df["耦合度C"] = df.apply(lambda r: coupling_degree(r["老龄化指数"], r["医疗资源指数"]), axis=1)
    df["资源倾斜优先级"] = df.apply(priority_label, axis=1)
    df = df.sort_values("匹配差值", ascending=True).reset_index(drop=True)
    df["匹配差值排名"] = np.arange(1, len(df) + 1)
    df.insert(0, "医疗口径", model_name)
    return df


def summarize_match(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    return {
        "匹配类型汇总": df["匹配类型"].value_counts().rename_axis("匹配类型").reset_index(name="省份数量"),
        "四象限汇总": df["四象限类型"].value_counts().rename_axis("四象限类型").reset_index(name="省份数量"),
        "重点关注地区": df[df["资源倾斜优先级"].isin(["第一优先级", "第二优先级"])].copy(),
        "高老龄低资源地区": df[df["四象限类型"] == "高老龄-低资源"].copy(),
    }


def save_match_workbook(df: pd.DataFrame, output_file: str | Path) -> None:
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    summaries = summarize_match(df)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="匹配分析结果", index=False)
        for sheet, table in summaries.items():
            table.to_excel(writer, sheet_name=sheet[:31], index=False)


def robustness_tables(results: dict[str, pd.DataFrame], base_model: str = "人口口径") -> dict[str, pd.DataFrame]:
    def quartile_shortage(series: pd.Series) -> pd.Series:
        rank_pct = series.rank(method="first", ascending=True) / len(series)
        return rank_pct <= 0.50

    enriched = {}
    for name, df in results.items():
        temp = df.copy()
        temp["固定阈值_是否资源不足"] = temp["匹配类型"].isin(["严重失配", "相对失配"])
        temp["四分位_是否资源不足"] = quartile_shortage(temp["匹配差值"])
        temp["阈值分类是否稳定"] = temp["固定阈值_是否资源不足"] == temp["四分位_是否资源不足"]
        temp["中位数四象限"] = temp.apply(
            lambda r: quadrant_type(r["老龄化指数"], r["医疗资源指数"], temp["老龄化指数"].median(), temp["医疗资源指数"].median()),
            axis=1,
        )
        temp["四象限是否变化"] = temp["四象限类型"] != temp["中位数四象限"]
        enriched[name] = temp

    compare = None
    for name, df in enriched.items():
        cols = ["地区", "老龄化指数", "医疗资源指数", "匹配差值", "固定阈值_是否资源不足", "四象限类型"]
        temp = df[cols].rename(columns={
            "医疗资源指数": f"{name}_医疗资源指数",
            "匹配差值": f"{name}_匹配差值",
            "固定阈值_是否资源不足": f"{name}_是否资源不足",
            "四象限类型": f"{name}_四象限类型",
        })
        if compare is None:
            compare = temp
        else:
            temp = temp.drop(columns=["老龄化指数"], errors="ignore")
            compare = pd.merge(compare, temp, on="地区", how="outer")

    model_names = list(results)
    gap_cols = [f"{name}_匹配差值" for name in model_names]
    shortage_cols = [f"{name}_是否资源不足" for name in model_names]
    compare["三口径_资源不足次数"] = compare[shortage_cols].sum(axis=1)
    compare["三口径_平均匹配差值"] = compare[gap_cols].mean(axis=1)
    compare = compare.sort_values(["三口径_资源不足次数", "三口径_平均匹配差值"], ascending=[False, True])

    corr = pd.DataFrame(index=model_names, columns=model_names, dtype=float)
    for a in model_names:
        for b in model_names:
            merged = pd.merge(results[a][["地区", "匹配差值"]], results[b][["地区", "匹配差值"]], on="地区", suffixes=("_a", "_b"))
            corr.loc[a, b] = merged["匹配差值_a"].corr(merged["匹配差值_b"], method="spearman")

    base = results[base_model] if base_model in results else next(iter(results.values()))
    base_set = set(base.loc[base["匹配类型"].isin(["严重失配", "相对失配"]), "地区"])
    overlap_rows = []
    for name, df in results.items():
        test_set = set(df.loc[df["匹配类型"].isin(["严重失配", "相对失配"]), "地区"])
        union = base_set | test_set
        overlap_rows.append({
            "比较口径": name,
            "基准资源不足地区数": len(base_set),
            "比较组资源不足地区数": len(test_set),
            "交集数": len(base_set & test_set),
            "Jaccard重合系数": len(base_set & test_set) / len(union) if union else np.nan,
        })

    return {
        "三口径对比总表": compare,
        "Spearman排序相关矩阵": corr,
        "重点地区重合率": pd.DataFrame(overlap_rows),
        **{f"{name}_完整稳健性": df for name, df in enriched.items()},
    }
