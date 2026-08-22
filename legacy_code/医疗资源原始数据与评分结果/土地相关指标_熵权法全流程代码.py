import math
import numpy as np
import pandas as pd

input_file = r"/mnt/data/2024医疗资源配置_Zhou2025_公平效率评价结果(1).xlsx"
sheet_name = "Data_2024_with_area"
output_file = r"/mnt/data/2024土地相关指标_熵权法评价结果_复现版.xlsx"

indicators = [
    "每平方千米医疗卫生机构数",
    "每平方千米床位数",
    "每平方千米卫生技术人员数",
    "每平方千米执业(助理)医师数",
    "每平方千米注册护士数"
]

df = pd.read_excel(input_file, sheet_name=sheet_name)
data = df[["地区", "四大区域"] + indicators].copy()

X = data[indicators].astype(float).to_numpy()
mins = X.min(axis=0)
maxs = X.max(axis=0)

# 功效系数法标准化（全部为正向指标）
Z = 0.1 + 0.9 * (X - mins) / (maxs - mins)

# 比重矩阵
P = Z / Z.sum(axis=0)

# 熵值
n = X.shape[0]
k = 1 / math.log(n)
with np.errstate(divide="ignore", invalid="ignore"):
    E = -k * np.where(P > 0, P * np.log(P), 0.0).sum(axis=0)

# 差异系数与权重
D = 1 - E
W = D / D.sum()

# 综合得分
S = Z.dot(W)

# 指标权重表
weight_df = pd.DataFrame({
    "指标": indicators,
    "最小值": mins,
    "最大值": maxs,
    "熵值_ej": E,
    "差异系数_dj": D,
    "权重_wj": W
})

# 省级结果
std_cols = [f"{c}_标准化" for c in indicators]
score_df = data.copy()
score_df[std_cols] = Z
score_df["综合得分"] = S
score_df = score_df.sort_values("综合得分", ascending=False).reset_index(drop=True)
score_df["排名"] = score_df.index + 1

# 区域汇总
region_df = score_df.groupby("四大区域")[indicators + ["综合得分"]].mean().reset_index()
region_df = region_df.sort_values("综合得分", ascending=False).reset_index(drop=True)
region_df["排名"] = region_df.index + 1

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    weight_df.to_excel(writer, sheet_name="指标权重", index=False)
    score_df.to_excel(writer, sheet_name="省级结果", index=False)
    region_df.to_excel(writer, sheet_name="区域汇总", index=False)

print("指标权重：")
print(weight_df)
print("\n省级结果前10：")
print(score_df.head(10)[["地区", "四大区域", "综合得分", "排名"]])
print("\n区域结果：")
print(region_df[["四大区域", "综合得分", "排名"]])
print(f"\n结果已保存到：{output_file}")
