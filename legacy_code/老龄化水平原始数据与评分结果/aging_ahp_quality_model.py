
import pandas as pd
import numpy as np
import math
from pathlib import Path

# =========================
# 1. 文件路径与工作表
# =========================
FILE_PATH = r"D:\文献\老龄化指标提取_中文最终版_扩展指标版.xlsx"
SHEET_NAME = "老龄化扩展候选指标_2023"
OUTPUT_DIR = Path(FILE_PATH).parent

# =========================
# 2. AHP质量导向拓展模型
#    准则层：
#    C1 年龄结构质量
#    C2 赡养负担质量
#    C3 人口再生产质量
#
#    指标层：
#    年龄结构质量：65岁及以上人口占比、15-64岁人口占比、0-14岁人口占比
#    赡养负担质量：老年抚养比、老少比、总抚养比
#    人口再生产质量：出生率、自然增长率
# =========================

DIRECTIONS = {
    "65岁及以上人口占比": "positive",
    "15-64岁人口占比": "negative",
    "0-14岁人口占比": "negative",
    "老年抚养比": "positive",
    "老少比": "positive",
    "总抚养比": "positive",
    "出生率_‰": "negative",
    "自然增长率_‰": "negative",
}

# 目标层 -> 准则层
A_CRITERIA = np.array([
    [1,   2,   4],
    [1/2, 1,   2],
    [1/4, 1/2, 1]
], dtype=float)

# 年龄结构质量：65岁及以上 > 15-64岁 > 0-14岁
A_AGE = np.array([
    [1,   2,   4],
    [1/2, 1,   2],
    [1/4, 1/2, 1]
], dtype=float)

# 赡养负担质量：老年抚养比 > 老少比 > 总抚养比
A_DEP = np.array([
    [1,   2,   3],
    [1/2, 1,   2],
    [1/3, 1/2, 1]
], dtype=float)

# 人口再生产质量：出生率 > 自然增长率
A_REP = np.array([
    [1,   2],
    [1/2, 1]
], dtype=float)

RI_TABLE = {
    1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12,
    6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49
}


def ahp_weights(matrix: np.ndarray):
    """
    根据判断矩阵计算AHP权重、最大特征值、一致性指标CI和一致性比率CR
    """
    eigvals, eigvecs = np.linalg.eig(matrix)
    max_idx = np.argmax(eigvals.real)
    lambda_max = float(eigvals.real[max_idx])
    w = np.abs(eigvecs[:, max_idx].real)
    w = w / w.sum()

    n = matrix.shape[0]
    ci = (lambda_max - n) / (n - 1) if n > 1 else 0.0
    ri = RI_TABLE[n]
    cr = ci / ri if ri != 0 else 0.0

    # 修正浮点误差
    if abs(cr) < 1e-12:
        cr = 0.0
    if abs(ci) < 1e-12:
        ci = 0.0
    if abs(lambda_max - n) < 1e-12:
        lambda_max = float(n)

    return w, lambda_max, ci, ri, cr


def min_max_standardize(df: pd.DataFrame, directions: dict) -> pd.DataFrame:
    """
    极差标准化
    正向指标： (x - min)/(max - min)
    逆向指标： (max - x)/(max - min)
    """
    result = pd.DataFrame(index=df.index)

    for col, direction in directions.items():
        x = pd.to_numeric(df[col], errors="coerce").astype(float)
        x_min = x.min()
        x_max = x.max()

        if math.isclose(x_max, x_min):
            result[col] = 1.0
            continue

        if direction == "positive":
            z = (x - x_min) / (x_max - x_min)
        elif direction == "negative":
            z = (x_max - x) / (x_max - x_min)
        else:
            raise ValueError(f"指标 {col} 的方向必须是 positive 或 negative")

        result[col] = z

    return result


def main():
    print("开始读取数据...")
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=3)

    needed_cols = ["地区", "年份"] + list(DIRECTIONS.keys())
    df = df[needed_cols].copy()

    # AHP权重
    w_criteria, lam1, ci1, ri1, cr1 = ahp_weights(A_CRITERIA)
    w_age, lam2, ci2, ri2, cr2 = ahp_weights(A_AGE)
    w_dep, lam3, ci3, ri3, cr3 = ahp_weights(A_DEP)
    w_rep, lam4, ci4, ri4, cr4 = ahp_weights(A_REP)

    # 计算全局权重
    global_weights = pd.Series({
        "65岁及以上人口占比": w_criteria[0] * w_age[0],
        "15-64岁人口占比": w_criteria[0] * w_age[1],
        "0-14岁人口占比": w_criteria[0] * w_age[2],
        "老年抚养比": w_criteria[1] * w_dep[0],
        "老少比": w_criteria[1] * w_dep[1],
        "总抚养比": w_criteria[1] * w_dep[2],
        "出生率_‰": w_criteria[2] * w_rep[0],
        "自然增长率_‰": w_criteria[2] * w_rep[1],
    })

    # 标准化
    std_df = min_max_standardize(df, DIRECTIONS)

    # 综合得分
    df["AHP老龄化质量得分"] = std_df.mul(global_weights, axis=1).sum(axis=1)

    # 排名
    result = df[["地区", "年份", "AHP老龄化质量得分"]].copy()
    result = result.sort_values(by="AHP老龄化质量得分", ascending=False).reset_index(drop=True)
    result["排名"] = result.index + 1
    result = result[["排名", "地区", "年份", "AHP老龄化质量得分"]]

    # 输出权重表
    criteria_weight_df = pd.DataFrame({
        "准则层": ["年龄结构质量", "赡养负担质量", "人口再生产质量"],
        "权重": w_criteria
    })

    local_weight_df = pd.DataFrame([
        ("年龄结构质量", "65岁及以上人口占比", w_age[0], global_weights["65岁及以上人口占比"]),
        ("年龄结构质量", "15-64岁人口占比", w_age[1], global_weights["15-64岁人口占比"]),
        ("年龄结构质量", "0-14岁人口占比", w_age[2], global_weights["0-14岁人口占比"]),
        ("赡养负担质量", "老年抚养比", w_dep[0], global_weights["老年抚养比"]),
        ("赡养负担质量", "老少比", w_dep[1], global_weights["老少比"]),
        ("赡养负担质量", "总抚养比", w_dep[2], global_weights["总抚养比"]),
        ("人口再生产质量", "出生率_‰", w_rep[0], global_weights["出生率_‰"]),
        ("人口再生产质量", "自然增长率_‰", w_rep[1], global_weights["自然增长率_‰"]),
    ], columns=["准则层", "指标层", "局部权重", "全局权重"]).sort_values(by="全局权重", ascending=False)

    consistency_df = pd.DataFrame([
        ("目标层-准则层", 3, lam1, ci1, ri1, cr1),
        ("年龄结构质量", 3, lam2, ci2, ri2, cr2),
        ("赡养负担质量", 3, lam3, ci3, ri3, cr3),
        ("人口再生产质量", 2, lam4, ci4, ri4, cr4),
    ], columns=["判断矩阵", "阶数n", "λmax", "CI", "RI", "CR"])
    consistency_df["一致性结论"] = consistency_df["CR"].apply(lambda x: "通过" if x < 0.10 else "未通过")

    output_file = OUTPUT_DIR / "AHP老龄化质量模型结果.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="综合得分排名", index=False)
        criteria_weight_df.to_excel(writer, sheet_name="准则层权重", index=False)
        local_weight_df.to_excel(writer, sheet_name="指标层权重", index=False)
        consistency_df.to_excel(writer, sheet_name="一致性检验", index=False)
        std_df_out = df[["地区", "年份"]].join(std_df)
        std_df_out.to_excel(writer, sheet_name="标准化数据", index=False)
        df.to_excel(writer, sheet_name="清洗后原始数据", index=False)

    print("\n========== AHP计算完成 ==========")
    print("\n【准则层权重】")
    print(criteria_weight_df)
    print("\n【指标层全局权重】")
    print(local_weight_df)
    print("\n【一致性检验】")
    print(consistency_df)
    print("\n【综合得分前10名】")
    print(result.head(10))
    print(f"\n结果已保存到：{output_file}")


if __name__ == "__main__":
    main()
