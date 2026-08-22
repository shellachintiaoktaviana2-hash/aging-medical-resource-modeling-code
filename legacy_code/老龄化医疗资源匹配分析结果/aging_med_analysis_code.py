# -*- coding: utf-8 -*-
"""
老龄化指数与医疗资源数量、空间结构、质量三层对应关系分析
输入：
  1. 老龄化熵权法结果(4).xlsx
  2. 医疗资源配置_AHP权重与结果_2024_质量侧重版(3).xlsx
输出：
  1. 老龄化_医疗资源三维对应关系分析结果.xlsx
  2. 四象限_数量.png、四象限_空间结构.png、四象限_质量.png
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import pearsonr, spearmanr

# ===================== 1. 路径设置 =====================
AGING_FILE = r"/mnt/data/老龄化熵权法结果(4).xlsx"
MEDICAL_FILE = r"/mnt/data/医疗资源配置_AHP权重与结果_2024_质量侧重版(3).xlsx"
OUT_EXCEL = r"/mnt/data/老龄化_医疗资源三维对应关系分析结果.xlsx"
OUT_DIR = r"/mnt/data/aging_med_analysis_plots"
os.makedirs(OUT_DIR, exist_ok=True)

# 若在本地电脑运行，请把上面路径改成自己的绝对路径，例如：
# AGING_FILE = r"D:\文献\老龄化熵权法结果(4).xlsx"
# MEDICAL_FILE = r"D:\文献\医疗资源配置_AHP权重与结果_2024_质量侧重版(3).xlsx"
# OUT_EXCEL = r"D:\文献\老龄化_医疗资源三维对应关系分析结果.xlsx"
# OUT_DIR = r"D:\文献\aging_med_analysis_plots"

# ===================== 2. 读取数据 =====================
aging = pd.read_excel(AGING_FILE, sheet_name="综合得分排名")
aging = aging[["地区", "年份", "综合得分", "排名"]].rename(
    columns={"年份": "老龄化年份", "综合得分": "U1_老龄化", "排名": "老龄化排名"}
)

medical = pd.read_excel(MEDICAL_FILE, sheet_name="省级结果", header=2)
weights = pd.read_excel(MEDICAL_FILE, sheet_name="最终权重", header=1)
weights.columns = ["指标", "原始数据列", "所属准则层", "准则层权重", "局部权重", "全局权重"]
weights = weights[weights["指标"] != "指标"].copy()
weights["全局权重"] = pd.to_numeric(weights["全局权重"])

# ===================== 3. 三个医疗资源维度的指标归类 =====================
DIMENSION_MAP = {
    "数量": [
        "C1_每千人口执业(助理)医师数",
        "C2_每万人口全科医生数",
    ],
    "空间结构": [
        "C7_每十万人三级医院数",
        "C8_每十万人三级甲等医院数",
    ],
    "质量": [
        "C3_全科医生占执业(助理)医师比",
        "C4_三级医院占医院比",
        "C5_三级甲等医院占医院比",
        "C6_三级甲等占三级医院比",
        "C9_信息化高配综合指数",
    ],
}

# 维度内权重：使用原AHP全局权重在该维度内重新归一化
map_rows = []
for dim, cols in DIMENSION_MAP.items():
    temp = weights.set_index("指标").loc[cols, ["全局权重"]].copy()
    temp["维度"] = dim
    temp["维度内归一化权重"] = temp["全局权重"] / temp["全局权重"].sum()
    temp = temp.reset_index()
    map_rows.append(temp)
map_df = pd.concat(map_rows, ignore_index=True)

# ===================== 4. 计算维度得分、Gap、耦合度、四象限 =====================
def classify_gap(gap: float) -> str:
    """根据Gap判断供需状态。阈值可按比赛需要调整。"""
    if gap < -0.05:
        return "短缺型"
    elif gap > 0.05:
        return "富余型"
    else:
        return "基本匹配"

result = aging.merge(
    medical[["地区"] + sum(DIMENSION_MAP.values(), []) + ["综合得分", "排名"]],
    on="地区",
    how="inner",
)
result = result.rename(columns={"综合得分": "U2_医疗综合得分", "排名": "医疗综合排名"})

for dim, cols in DIMENSION_MAP.items():
    dim_weight = map_df.loc[map_df["维度"] == dim, ["指标", "维度内归一化权重"]]
    dim_weight = dim_weight.set_index("指标")["维度内归一化权重"]

    # U2维度得分
    result[f"U2_{dim}"] = (result[cols] * dim_weight.loc[cols].values).sum(axis=1)

    # Gap = U2 - U1
    result[f"Gap_{dim}"] = result[f"U2_{dim}"] - result["U1_老龄化"]
    result[f"状态_{dim}"] = result[f"Gap_{dim}"].apply(classify_gap)

    # 耦合度 C = 2*sqrt(U1*U2)/(U1+U2)
    denominator = result["U1_老龄化"] + result[f"U2_{dim}"]
    result[f"Coupling_{dim}"] = np.where(
        denominator > 0,
        2 * np.sqrt(result["U1_老龄化"] * result[f"U2_{dim}"]) / denominator,
        0,
    )

    # 四象限划分：以U1均值与当前维度U2均值作为分界线
    x_mean = result["U1_老龄化"].mean()
    y_mean = result[f"U2_{dim}"].mean()

    def quadrant(row):
        x = row["U1_老龄化"]
        y = row[f"U2_{dim}"]
        if x >= x_mean and y >= y_mean:
            return "高老龄—高资源"
        elif x >= x_mean and y < y_mean:
            return "高老龄—低资源"
        elif x < x_mean and y >= y_mean:
            return "低老龄—高资源"
        else:
            return "低老龄—低资源"

    result[f"象限_{dim}"] = result.apply(quadrant, axis=1)

# ===================== 5. 补充相关性分析 =====================
corr_rows = []
for dim in DIMENSION_MAP.keys():
    pearson_r, pearson_p = pearsonr(result["U1_老龄化"], result[f"U2_{dim}"])
    spearman_rho, spearman_p = spearmanr(result["U1_老龄化"], result[f"U2_{dim}"])
    corr_rows.append({
        "维度": dim,
        "Pearson_r": pearson_r,
        "Pearson_p": pearson_p,
        "Spearman_rho": spearman_rho,
        "Spearman_p": spearman_p,
        "U1均值": result["U1_老龄化"].mean(),
        "U2均值": result[f"U2_{dim}"].mean(),
        "平均Gap": result[f"Gap_{dim}"].mean(),
        "平均耦合度": result[f"Coupling_{dim}"].mean(),
        "短缺型个数": (result[f"状态_{dim}"] == "短缺型").sum(),
        "基本匹配个数": (result[f"状态_{dim}"] == "基本匹配").sum(),
        "富余型个数": (result[f"状态_{dim}"] == "富余型").sum(),
    })
corr_df = pd.DataFrame(corr_rows)

# ===================== 6. 输出Excel结果 =====================
with pd.ExcelWriter(OUT_EXCEL, engine="openpyxl") as writer:
    result.sort_values("老龄化排名").to_excel(writer, sheet_name="总表", index=False)
    map_df.to_excel(writer, sheet_name="维度指标归类与权重", index=False)
    corr_df.to_excel(writer, sheet_name="补充相关性", index=False)
    for dim in DIMENSION_MAP.keys():
        cols = ["地区", "U1_老龄化", f"U2_{dim}", f"Gap_{dim}", f"状态_{dim}", f"Coupling_{dim}", f"象限_{dim}"]
        result[cols].sort_values(f"Gap_{dim}").to_excel(writer, sheet_name=f"{dim}层分析", index=False)

# ===================== 7. 绘制四象限图 =====================
plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "Noto Sans CJK JP", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

for dim in DIMENSION_MAP.keys():
    fig, ax = plt.subplots(figsize=(10, 8))
    x = result["U1_老龄化"]
    y = result[f"U2_{dim}"]

    ax.scatter(x, y)
    ax.axvline(x.mean(), linestyle="--")
    ax.axhline(y.mean(), linestyle="--")

    for _, row in result.iterrows():
        ax.text(row["U1_老龄化"] + 0.005, row[f"U2_{dim}"] + 0.005, row["地区"], fontsize=8)

    ax.set_xlabel("老龄化指数 U1")
    ax.set_ylabel(f"{dim}医疗资源指数 U2")
    ax.set_title(f"老龄化与{dim}医疗资源四象限图")
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"四象限_{dim}.png"), dpi=200)
    plt.close(fig)

print("分析完成")
print("Excel结果：", OUT_EXCEL)
print("图片文件夹：", OUT_DIR)
