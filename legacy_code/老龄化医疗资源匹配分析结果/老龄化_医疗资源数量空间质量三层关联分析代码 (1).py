# -*- coding: utf-8 -*-
"""
老龄化—医疗资源数量/空间结构/质量三层关联分析
方法：Gap匹配差值 + 耦合度 + 四象限
运行方式：将本脚本与3个Excel放在同一文件夹，或修改下方BASE_DIR与文件名后运行。
"""

import os
import math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ========== 1. 路径设置 ==========
BASE_DIR = os.getcwd()

AGING_FILE = os.path.join(BASE_DIR, "老龄化熵权法结果(4).xlsx")
QUANTITY_FILE = os.path.join(BASE_DIR, "U1等权_每千人口熵权匹配分析结果(3).xlsx")
SPACE_FILE = os.path.join(BASE_DIR, "2024土地相关指标_熵权法评价结果(2).xlsx")

OUTPUT_XLSX = os.path.join(BASE_DIR, "老龄化_医疗资源数量空间质量三层关联分析结果.xlsx")
OUTPUT_GAP_HEATMAP = os.path.join(BASE_DIR, "三层匹配差值Gap热力图.png")


# ========== 2. 基础函数 ==========
def minmax_standardize(s: pd.Series) -> pd.Series:
    """极差标准化：U'=(U-min)/(max-min)"""
    s = s.astype(float)
    if s.max() == s.min():
        return pd.Series(np.zeros(len(s)), index=s.index)
    return (s - s.min()) / (s.max() - s.min())


def entropy_weight(df_std: pd.DataFrame):
    """
    熵权法。
    输入：已经0—1标准化后的正向指标矩阵。
    输出：权重Series、熵值Series、差异系数Series。
    """
    X = df_std.astype(float).copy()
    n = X.shape[0]
    eps = 1e-12
    P = X / (X.sum(axis=0) + eps)
    E = -(P * np.log(P + eps)).sum(axis=0) / np.log(n)
    D = 1 - E
    W = D / D.sum()
    return W, E, D


def gap_grade(g):
    if g < -0.20:
        return "严重短缺型"
    elif g < -0.05:
        return "相对短缺型"
    elif g <= 0.05:
        return "基本匹配型"
    elif g <= 0.20:
        return "相对富余型"
    else:
        return "明显富余型"


def gap_direction(g):
    if g < -0.05:
        return "资源滞后"
    elif g <= 0.05:
        return "基本匹配"
    else:
        return "资源富余"


def coupling_degree(u1, u2):
    """C=2*sqrt(U1*U2)/(U1+U2)"""
    if u1 + u2 == 0:
        return 0
    return 2 * math.sqrt(max(0, u1 * u2)) / (u1 + u2)


def coupling_grade(c):
    if c < 0.50:
        return "低耦合"
    elif c < 0.70:
        return "中等耦合"
    elif c < 0.90:
        return "较高耦合"
    else:
        return "高耦合"


def quadrant(u1, u2, u1_mean, u2_mean):
    if u1 >= u1_mean and u2 >= u2_mean:
        return "高老龄—高资源"
    elif u1 >= u1_mean and u2 < u2_mean:
        return "高老龄—低资源"
    elif u1 < u1_mean and u2 >= u2_mean:
        return "低老龄—高资源"
    else:
        return "低老龄—低资源"


def run_layer_analysis(base_df, layer_name, u2_col):
    """对单个医疗资源层面进行Gap、耦合度和四象限分析。"""
    df = base_df[["地区", "四大区域", "年份", "U1老龄化原始得分", u2_col]].copy()
    df = df.rename(columns={u2_col: "U2医疗资源原始得分"})
    df["层面"] = layer_name

    df["U1标准化"] = minmax_standardize(df["U1老龄化原始得分"])
    df["U2标准化"] = minmax_standardize(df["U2医疗资源原始得分"])

    u1_mean = df["U1标准化"].mean()
    u2_mean = df["U2标准化"].mean()

    df["Gap匹配差值"] = df["U2标准化"] - df["U1标准化"]
    df["Gap方向"] = df["Gap匹配差值"].apply(gap_direction)
    df["Gap等级"] = df["Gap匹配差值"].apply(gap_grade)
    df["耦合度C"] = df.apply(lambda r: coupling_degree(r["U1标准化"], r["U2标准化"]), axis=1)
    df["耦合等级"] = df["耦合度C"].apply(coupling_grade)
    df["四象限类型"] = df.apply(lambda r: quadrant(r["U1标准化"], r["U2标准化"], u1_mean, u2_mean), axis=1)
    df = df.sort_values("Gap匹配差值", ascending=True).reset_index(drop=True)
    df["短缺优先序"] = np.arange(1, len(df) + 1)
    return df


# ========== 3. 读取数据 ==========
aging = pd.read_excel(AGING_FILE, sheet_name="综合得分排名")
aging = aging.rename(columns={"综合得分": "U1老龄化原始得分"})

quantity = pd.read_excel(QUANTITY_FILE, sheet_name="每千人口熵权得分排序")
quantity = quantity.rename(columns={"每千人口熵权得分": "U2数量原始得分"})

space = pd.read_excel(SPACE_FILE, sheet_name="省级结果")
space = space.rename(columns={"综合得分": "U2空间结构原始得分"})

# ========== 4. 构造质量代理得分 ==========
# 说明：如果已有独立的医疗质量指标，可直接替换此部分。
# 当前根据上传数据中的床位密度与人力密度构造“人力—床位配置结构”质量代理指标。
space["每床卫生技术人员数"] = space["每平方千米卫生技术人员数"] / space["每平方千米床位数"]
space["每床执业医师数"] = space["每平方千米执业(助理)医师数"] / space["每平方千米床位数"]
space["每床注册护士数"] = space["每平方千米注册护士数"] / space["每平方千米床位数"]
space["护士_医师比"] = space["每平方千米注册护士数"] / space["每平方千米执业(助理)医师数"]

quality_cols = ["每床卫生技术人员数", "每床执业医师数", "每床注册护士数", "护士_医师比"]
quality_std = space[quality_cols].apply(minmax_standardize)
quality_weight, quality_entropy, quality_diff = entropy_weight(quality_std)
space["U2质量代理原始得分"] = (quality_std * quality_weight).sum(axis=1)
space["质量代理排名"] = space["U2质量代理原始得分"].rank(ascending=False, method="min").astype(int)

quality_weight_table = pd.DataFrame({
    "质量代理指标": quality_cols,
    "熵值E": quality_entropy.values,
    "差异系数d": quality_diff.values,
    "权重w": quality_weight.values
})

# ========== 5. 合并三个层面数据 ==========
base = aging[["地区", "年份", "U1老龄化原始得分"]].merge(
    space[["地区", "四大区域", "U2空间结构原始得分", "U2质量代理原始得分"]],
    on="地区",
    how="inner"
).merge(
    quantity[["地区", "U2数量原始得分"]],
    on="地区",
    how="inner"
)

quantity_result = run_layer_analysis(base, "数量层面", "U2数量原始得分")
space_result = run_layer_analysis(base, "空间结构层面", "U2空间结构原始得分")
quality_result = run_layer_analysis(base, "质量代理层面", "U2质量代理原始得分")
long_result = pd.concat([quantity_result, space_result, quality_result], ignore_index=True)

# ========== 6. 汇总表 ==========
summary_list = []
for layer, sub in long_result.groupby("层面"):
    counts = sub["Gap等级"].value_counts()
    q_counts = sub["四象限类型"].value_counts()
    summary_list.append({
        "层面": layer,
        "Pearson相关": sub["U1老龄化原始得分"].corr(sub["U2医疗资源原始得分"], method="pearson"),
        "Spearman相关": sub["U1老龄化原始得分"].corr(sub["U2医疗资源原始得分"], method="spearman"),
        "Gap均值": sub["Gap匹配差值"].mean(),
        "严重短缺型数量": counts.get("严重短缺型", 0),
        "相对短缺型数量": counts.get("相对短缺型", 0),
        "基本匹配型数量": counts.get("基本匹配型", 0),
        "相对富余型数量": counts.get("相对富余型", 0),
        "明显富余型数量": counts.get("明显富余型", 0),
        "高老龄—低资源数量": q_counts.get("高老龄—低资源", 0),
        "高老龄—低资源地区": "、".join(sub.loc[sub["四象限类型"] == "高老龄—低资源", "地区"]),
        "严重短缺地区": "、".join(sub.loc[sub["Gap等级"] == "严重短缺型", "地区"]),
        "相对短缺地区": "、".join(sub.loc[sub["Gap等级"] == "相对短缺型", "地区"])
    })

summary = pd.DataFrame(summary_list)

# 宽表：便于比较每个省在三个层面的短板
wide = base[["地区", "四大区域", "年份", "U1老龄化原始得分"]].copy()
for layer, sub in long_result.groupby("层面"):
    prefix = {"数量层面": "数量", "空间结构层面": "空间", "质量代理层面": "质量"}[layer]
    tmp = sub[["地区", "U2医疗资源原始得分", "U2标准化", "Gap匹配差值", "Gap等级", "耦合度C", "四象限类型", "短缺优先序"]].copy()
    tmp.columns = ["地区", f"U2{prefix}原始得分", f"U2{prefix}标准化", f"{prefix}Gap", f"{prefix}Gap等级", f"{prefix}耦合度C", f"{prefix}四象限", f"{prefix}短缺优先序"]
    wide = wide.merge(tmp, on="地区", how="left")

gap_cols = ["数量Gap", "空间Gap", "质量Gap"]
wide["短板层面数量"] = (wide[gap_cols] < -0.05).sum(axis=1)
wide["高老龄低资源层面数量"] = (
    (wide["数量四象限"] == "高老龄—低资源").astype(int)
    + (wide["空间四象限"] == "高老龄—低资源").astype(int)
    + (wide["质量四象限"] == "高老龄—低资源").astype(int)
)
wide["综合短缺识别"] = np.where(
    (wide["短板层面数量"] >= 2) | (wide["高老龄低资源层面数量"] >= 2),
    "重点补短板",
    np.where(wide["短板层面数量"] == 1, "单项短板", "非短缺主导")
)

# ========== 7. 输出Excel ==========
with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="层面汇总", index=False)
    long_result.to_excel(writer, sheet_name="三层长表", index=False)
    wide.to_excel(writer, sheet_name="省级宽表", index=False)
    quality_weight_table.to_excel(writer, sheet_name="质量代理权重", index=False)
    space[["地区", "四大区域"] + quality_cols + ["U2质量代理原始得分", "质量代理排名"]].to_excel(
        writer, sheet_name="质量代理指标计算", index=False
    )

# ========== 8. 绘制Gap热力图 ==========
heat = wide[["地区"] + gap_cols].copy()
heat["最小Gap"] = heat[gap_cols].min(axis=1)
heat = heat.sort_values("最小Gap", ascending=True)

plt.figure(figsize=(7, 10), dpi=180)
plt.imshow(heat[gap_cols], aspect="auto", cmap="coolwarm", vmin=-1, vmax=1)
plt.xticks(range(3), gap_cols, rotation=15)
plt.yticks(range(len(heat)), heat["地区"], fontsize=8)
plt.colorbar(label="Gap=U2'-U1'")
plt.title("三层匹配差值Gap热力图")
for i in range(len(heat)):
    for j, col in enumerate(gap_cols):
        plt.text(j, i, f"{heat.iloc[i][col]:.2f}", ha="center", va="center", fontsize=6)
plt.tight_layout()
plt.savefig(OUTPUT_GAP_HEATMAP, bbox_inches="tight")
plt.close()

print("分析完成：")
print(OUTPUT_XLSX)
print(OUTPUT_GAP_HEATMAP)
