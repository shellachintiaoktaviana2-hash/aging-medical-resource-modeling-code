from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from .utils import efficacy_standardize, ensure_dir, minmax_standardize, normalize_province_column


RI_TABLE = {1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}


def ahp_weights(matrix: np.ndarray) -> tuple[np.ndarray, float, float, float, float]:
    A = np.array(matrix, dtype=float)
    eigvals, eigvecs = np.linalg.eig(A)
    idx = np.argmax(eigvals.real)
    lam = float(eigvals.real[idx])
    weights = np.abs(eigvecs[:, idx].real)
    weights = weights / weights.sum()
    n = A.shape[0]
    ci = (lam - n) / (n - 1) if n > 1 else 0.0
    ri = RI_TABLE[n]
    cr = ci / ri if ri else 0.0
    if abs(ci) < 1e-12:
        ci = 0.0
    if abs(cr) < 1e-12:
        cr = 0.0
    if abs(lam - n) < 1e-12:
        lam = float(n)
    return weights, lam, ci, ri, cr


AGING_DIRECTIONS = {
    "65岁及以上人口占比": "positive",
    "15-64岁人口占比": "negative",
    "0-14岁人口占比": "negative",
    "老年抚养比": "positive",
    "老少比": "positive",
    "总抚养比": "positive",
    "出生率_‰": "negative",
    "自然增长率_‰": "negative",
}


def aging_ahp_weights() -> tuple[pd.Series, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    criteria = np.array([[1, 2, 4], [1 / 2, 1, 2], [1 / 4, 1 / 2, 1]], dtype=float)
    age = np.array([[1, 2, 4], [1 / 2, 1, 2], [1 / 4, 1 / 2, 1]], dtype=float)
    dep = np.array([[1, 2, 3], [1 / 2, 1, 2], [1 / 3, 1 / 2, 1]], dtype=float)
    rep = np.array([[1, 2], [1 / 2, 1]], dtype=float)

    wc, lam1, ci1, ri1, cr1 = ahp_weights(criteria)
    wa, lam2, ci2, ri2, cr2 = ahp_weights(age)
    wd, lam3, ci3, ri3, cr3 = ahp_weights(dep)
    wr, lam4, ci4, ri4, cr4 = ahp_weights(rep)

    global_w = pd.Series({
        "65岁及以上人口占比": wc[0] * wa[0],
        "15-64岁人口占比": wc[0] * wa[1],
        "0-14岁人口占比": wc[0] * wa[2],
        "老年抚养比": wc[1] * wd[0],
        "老少比": wc[1] * wd[1],
        "总抚养比": wc[1] * wd[2],
        "出生率_‰": wc[2] * wr[0],
        "自然增长率_‰": wc[2] * wr[1],
    })

    criteria_df = pd.DataFrame({"准则层": ["年龄结构质量", "赡养负担质量", "人口再生产质量"], "权重": wc})
    local_df = pd.DataFrame([
        ("年龄结构质量", "65岁及以上人口占比", wa[0], global_w["65岁及以上人口占比"]),
        ("年龄结构质量", "15-64岁人口占比", wa[1], global_w["15-64岁人口占比"]),
        ("年龄结构质量", "0-14岁人口占比", wa[2], global_w["0-14岁人口占比"]),
        ("赡养负担质量", "老年抚养比", wd[0], global_w["老年抚养比"]),
        ("赡养负担质量", "老少比", wd[1], global_w["老少比"]),
        ("赡养负担质量", "总抚养比", wd[2], global_w["总抚养比"]),
        ("人口再生产质量", "出生率_‰", wr[0], global_w["出生率_‰"]),
        ("人口再生产质量", "自然增长率_‰", wr[1], global_w["自然增长率_‰"]),
    ], columns=["准则层", "指标层", "局部权重", "全局权重"]).sort_values("全局权重", ascending=False)
    consistency_df = pd.DataFrame([
        ("目标层-准则层", 3, lam1, ci1, ri1, cr1),
        ("年龄结构质量", 3, lam2, ci2, ri2, cr2),
        ("赡养负担质量", 3, lam3, ci3, ri3, cr3),
        ("人口再生产质量", 2, lam4, ci4, ri4, cr4),
    ], columns=["判断矩阵", "阶数n", "lambda_max", "CI", "RI", "CR"])
    consistency_df["一致性结论"] = np.where(consistency_df["CR"] < 0.10, "通过", "未通过")
    return global_w, criteria_df, local_df, consistency_df


def run_aging_ahp(data: pd.DataFrame, output_file: str | Path) -> dict[str, pd.DataFrame]:
    df = normalize_province_column(data, "地区")
    needed = ["地区", "年份"] + list(AGING_DIRECTIONS.keys())
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"AHP老龄化输入缺少列：{missing}")

    raw = df[needed].copy()
    global_w, criteria_df, local_df, consistency_df = aging_ahp_weights()
    std = raw[["地区", "年份"]].copy()
    for col, direction in AGING_DIRECTIONS.items():
        std[col] = minmax_standardize(raw[col], direction=direction, fill_equal=1.0)

    result = raw[["地区", "年份"]].copy()
    result["AHP老龄化质量得分"] = std[list(AGING_DIRECTIONS.keys())].mul(global_w, axis=1).sum(axis=1)
    result["排名"] = result["AHP老龄化质量得分"].rank(ascending=False, method="min").astype(int)
    result = result.sort_values(["排名", "AHP老龄化质量得分"], ascending=[True, False]).reset_index(drop=True)
    result = result[["排名", "地区", "年份", "AHP老龄化质量得分"]]

    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="综合得分排名", index=False)
        criteria_df.to_excel(writer, sheet_name="准则层权重", index=False)
        local_df.to_excel(writer, sheet_name="指标层权重", index=False)
        consistency_df.to_excel(writer, sheet_name="一致性检验", index=False)
        std.to_excel(writer, sheet_name="标准化数据", index=False)
        raw.to_excel(writer, sheet_name="清洗后原始数据", index=False)

    return {"result": result, "criteria_weights": criteria_df, "indicator_weights": local_df, "consistency": consistency_df}


QUALITY_INDICATOR_MAP = {
    "C1_每千人口执业(助理)医师数": "每千人口执业(助理)医师数_计算",
    "C2_每万人口全科医生数": "每万人口全科医生数_计算",
    "C3_全科医生占执业(助理)医师比": "全科医生占执业(助理)医师比_%_计算",
    "C4_三级医院占医院比": "三级医院占医院比_%_计算",
    "C5_三级甲等医院占医院比": "三级甲等医院占医院比_%_计算",
    "C6_三级甲等占三级医院比": "三级甲等占三级医院比_%_计算",
    "C7_每十万人三级医院数": "每十万人三级医院数_计算",
    "C8_每十万人三级甲等医院数": "每十万人三级甲等医院数_计算",
    "C9_信息化高配综合指数": "信息化高配综合指数_计算",
}


def medical_quality_ahp_weights() -> pd.DataFrame:
    criteria = np.array([[1, 1 / 3, 1 / 2, 2], [3, 1, 2, 4], [2, 1 / 2, 1, 3], [1 / 2, 1 / 4, 1 / 3, 1]], dtype=float)
    b1 = np.array([[1, 1 / 3, 1 / 3], [3, 1, 1], [3, 1, 1]], dtype=float)
    b2 = np.array([[1, 1 / 4, 1 / 3], [4, 1, 2], [3, 1 / 2, 1]], dtype=float)
    b3 = np.array([[1, 1 / 3], [3, 1]], dtype=float)
    cw, *_ = ahp_weights(criteria)
    b1w, *_ = ahp_weights(b1)
    b2w, *_ = ahp_weights(b2)
    b3w, *_ = ahp_weights(b3)
    rows = [
        ("C1_每千人口执业(助理)医师数", "B1_人员质量与基层能力", cw[0], b1w[0]),
        ("C2_每万人口全科医生数", "B1_人员质量与基层能力", cw[0], b1w[1]),
        ("C3_全科医生占执业(助理)医师比", "B1_人员质量与基层能力", cw[0], b1w[2]),
        ("C4_三级医院占医院比", "B2_高等级优质医院水平", cw[1], b2w[0]),
        ("C5_三级甲等医院占医院比", "B2_高等级优质医院水平", cw[1], b2w[1]),
        ("C6_三级甲等占三级医院比", "B2_高等级优质医院水平", cw[1], b2w[2]),
        ("C7_每十万人三级医院数", "B3_优质资源可及性", cw[2], b3w[0]),
        ("C8_每十万人三级甲等医院数", "B3_优质资源可及性", cw[2], b3w[1]),
        ("C9_信息化高配综合指数", "B4_信息化质量支撑", cw[3], 1.0),
    ]
    out = pd.DataFrame(rows, columns=["指标", "所属准则层", "准则层权重", "局部权重"])
    out["全局权重"] = out["准则层权重"] * out["局部权重"]
    out["全局权重"] = out["全局权重"] / out["全局权重"].sum()
    out["原始数据列"] = out["指标"].map(QUALITY_INDICATOR_MAP)
    return out.sort_values("全局权重", ascending=False).reset_index(drop=True)


def run_medical_quality_ahp(data: pd.DataFrame, output_file: str | Path) -> dict[str, pd.DataFrame]:
    df = normalize_province_column(data, "地区")
    weight_df = medical_quality_ahp_weights()
    missing = [c for c in weight_df["原始数据列"] if c not in df.columns]
    if missing:
        raise ValueError(f"医疗质量AHP输入缺少列：{missing}")

    score = df[["地区"]].copy()
    for _, row in weight_df.iterrows():
        score[row["指标"]] = efficacy_standardize(df[row["原始数据列"]], direction="positive")
    score["综合得分"] = sum(score[row["指标"]] * row["全局权重"] for _, row in weight_df.iterrows())
    score["排名"] = score["综合得分"].rank(ascending=False, method="min").astype(int)
    score = score.sort_values(["排名", "综合得分"], ascending=[True, False]).reset_index(drop=True)

    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        weight_df.to_excel(writer, sheet_name="最终权重", index=False)
        score.to_excel(writer, sheet_name="省级结果", index=False)
    return {"weights": weight_df, "result": score}
