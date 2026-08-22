# -*- coding: utf-8 -*-
"""
老龄化指数 U1 与医疗资源配置指数 U2 的匹配差值分析
运行前请确保安装：
    pip install pandas openpyxl numpy

输入：
1. 老龄化熵权法结果(3).xlsx，使用工作表：综合得分排名
2. U1等权_每千人口熵权匹配分析结果(3).xlsx，使用工作表：每千人口熵权得分排序

输出：
老龄化_医疗资源匹配差值分析结果_复现版.xlsx
"""

import numpy as np
import pandas as pd


# ========== 1. 文件路径 ==========
# 如在本地运行，请改成自己的文件路径，例如 r"D:\文献\老龄化熵权法结果(3).xlsx"
U1_FILE = "老龄化熵权法结果(3).xlsx"
U2_FILE = "U1等权_每千人口熵权匹配分析结果(3).xlsx"
OUTPUT_FILE = "老龄化_医疗资源匹配差值分析结果_复现版.xlsx"


# ========== 2. 读取数据 ==========
u1 = pd.read_excel(U1_FILE, sheet_name="综合得分排名")
u2 = pd.read_excel(U2_FILE, sheet_name="每千人口熵权得分排序")

# U1：老龄化指数
u1 = u1[["地区", "年份", "综合得分", "排名"]].copy()
u1 = u1.rename(columns={
    "综合得分": "U1原始得分",
    "排名": "U1排名"
})

# U2：医疗资源配置指数
u2 = u2[["地区", "每千人口熵权得分", "每千人口排序"]].copy()
u2 = u2.rename(columns={
    "每千人口熵权得分": "U2原始得分",
    "每千人口排序": "U2排名"
})


# ========== 3. 按地区匹配 ==========
data = pd.merge(u1, u2, on="地区", how="inner")

if len(data) != len(u1) or len(data) != len(u2):
    missing_in_u2 = sorted(set(u1["地区"]) - set(u2["地区"]))
    missing_in_u1 = sorted(set(u2["地区"]) - set(u1["地区"]))
    print("警告：两张表地区未完全匹配。")
    print("U1中有但U2中没有：", missing_in_u2)
    print("U2中有但U1中没有：", missing_in_u1)


# ========== 4. 极差标准化 ==========
def minmax_standardize(series: pd.Series) -> pd.Series:
    """X'=(X-min)/(max-min)，若极差为0则返回0。"""
    min_v = series.min()
    max_v = series.max()
    if np.isclose(max_v - min_v, 0):
        return pd.Series(0, index=series.index)
    return (series - min_v) / (max_v - min_v)


data["U1标准化"] = minmax_standardize(data["U1原始得分"])
data["U2标准化"] = minmax_standardize(data["U2原始得分"])


# ========== 5. 计算匹配差值 ==========
# Mi = U2i - U1i
data["Mi=U2−U1"] = data["U2标准化"] - data["U1标准化"]
data["原始差值参考"] = data["U2原始得分"] - data["U1原始得分"]


# ========== 6. 按阈值分类 ==========
def classify_mi(mi: float) -> str:
    if mi < -0.20:
        return "严重失配"
    elif mi < -0.05:
        return "相对失配"
    elif mi <= 0.05:
        return "基本匹配"
    elif mi <= 0.20:
        return "相对超前"
    else:
        return "明显超前"


def explain_direction(mi: float) -> str:
    if mi < -0.05:
        return "医疗资源低于老龄化压力"
    elif mi <= 0.05:
        return "二者基本匹配"
    else:
        return "医疗资源相对超前"


data["匹配类型"] = data["Mi=U2−U1"].apply(classify_mi)
data["方向解释"] = data["Mi=U2−U1"].apply(explain_direction)


# ========== 7. 整理输出 ==========
data = data[[
    "地区", "年份",
    "U1原始得分", "U1标准化", "U1排名",
    "U2原始得分", "U2标准化", "U2排名",
    "Mi=U2−U1", "匹配类型", "方向解释", "原始差值参考"
]].copy()

data = data.sort_values("Mi=U2−U1", ascending=True).reset_index(drop=True)
data.insert(0, "序号", np.arange(1, len(data) + 1))

summary = (
    data.groupby("匹配类型", as_index=False)
    .agg(数量=("地区", "count"))
)

order = pd.DataFrame({
    "匹配类型": ["严重失配", "相对失配", "基本匹配", "相对超前", "明显超前"],
    "分类顺序": [1, 2, 3, 4, 5]
})
summary = order.merge(summary, on="匹配类型", how="left").fillna({"数量": 0})
summary["数量"] = summary["数量"].astype(int)
summary["占比"] = summary["数量"] / summary["数量"].sum()
summary = summary.drop(columns=["分类顺序"])

params = pd.DataFrame({
    "指标": ["U1老龄化指数", "U2医疗资源配置指数"],
    "最小值": [data["U1原始得分"].min(), data["U2原始得分"].min()],
    "最大值": [data["U1原始得分"].max(), data["U2原始得分"].max()],
    "标准化公式": [
        "U1'=(U1−min(U1))/(max(U1)−min(U1))",
        "U2'=(U2−min(U2))/(max(U2)−min(U2))"
    ]
})

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    data.to_excel(writer, sheet_name="匹配分析结果", index=False)
    summary.to_excel(writer, sheet_name="分类汇总", index=False)
    params.to_excel(writer, sheet_name="标准化参数", index=False)

print("分析完成，结果已输出：", OUTPUT_FILE)
print(summary)
