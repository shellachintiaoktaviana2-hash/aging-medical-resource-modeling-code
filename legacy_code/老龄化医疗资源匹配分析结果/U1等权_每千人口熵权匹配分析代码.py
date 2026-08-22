import pandas as pd
import numpy as np
from pathlib import Path

base = Path(__file__).resolve().parent
aging_path = base / "老龄化熵权法结果(2).xlsx"
med_path = base / "2024医疗资源配置_熵权TOPSIS_系统聚类结果(2).xlsx"
out_path = base / "U1等权_每千人口熵权匹配分析结果_代码运行输出.xlsx"

aging_std = pd.read_excel(aging_path, sheet_name="标准化矩阵")
inf_std = pd.read_excel(med_path, sheet_name="基础设施_标准化")
hr_std = pd.read_excel(med_path, sheet_name="人力资源_标准化")
svc_std = pd.read_excel(med_path, sheet_name="医疗服务_标准化")

aging_cols = [c for c in aging_std.columns if c not in ["地区", "年份"]]
aging_equal = aging_std[["地区", "年份"]].copy()
aging_equal["U1_原始系统得分(等权平均)"] = aging_std[aging_cols].mean(axis=1)
aging_equal["U1_再标准化"] = (
    aging_equal["U1_原始系统得分(等权平均)"] - aging_equal["U1_原始系统得分(等权平均)"].min()
) / (
    aging_equal["U1_原始系统得分(等权平均)"].max() - aging_equal["U1_原始系统得分(等权平均)"].min()
)
aging_equal["U1排序"] = aging_equal["U1_原始系统得分(等权平均)"].rank(ascending=False, method="min").astype(int)

med_all = inf_std.merge(hr_std, on="地区").merge(svc_std, on="地区")
med_cols = [c for c in med_all.columns if c != "地区"]
u2_df = med_all[["地区"]].copy()
u2_df["U2_原始系统得分(等权平均)"] = med_all[med_cols].mean(axis=1)
u2_df["U2_再标准化"] = (
    u2_df["U2_原始系统得分(等权平均)"] - u2_df["U2_原始系统得分(等权平均)"].min()
) / (
    u2_df["U2_原始系统得分(等权平均)"].max() - u2_df["U2_原始系统得分(等权平均)"].min()
)
u2_df["U2排序"] = u2_df["U2_原始系统得分(等权平均)"].rank(ascending=False, method="min").astype(int)

match_df = aging_equal[["地区", "年份", "U1_原始系统得分(等权平均)", "U1_再标准化", "U1排序"]].merge(
    u2_df[["地区", "U2_原始系统得分(等权平均)", "U2_再标准化", "U2排序"]],
    on="地区"
)
match_df["Mi=U2-U1"] = match_df["U2_再标准化"] - match_df["U1_再标准化"]

def classify(mi):
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

match_df["匹配类型"] = match_df["Mi=U2-U1"].apply(classify)
match_df["Mi排序(升序)"] = match_df["Mi=U2-U1"].rank(ascending=True, method="min").astype(int)
match_df = match_df.sort_values("Mi=U2-U1", ascending=True).reset_index(drop=True)

per1000 = inf_std.merge(hr_std, on="地区")
per_cols = [c for c in per1000.columns if c != "地区"]
X = per1000[per_cols].copy()
P = X.div(X.sum(axis=0), axis=1)
n = len(P)
k = 1 / np.log(n)

rows = []
for col in per_cols:
    p = P[col].astype(float).replace(0, np.nan)
    e = -k * np.nansum(p * np.log(p))
    d = 1 - e
    rows.append([col, e, d])

weights_df = pd.DataFrame(rows, columns=["指标", "熵值E", "差异系数d"])
weights_df["权重w"] = weights_df["差异系数d"] / weights_df["差异系数d"].sum()

weight_map = weights_df.set_index("指标")["权重w"]
per1000_score = per1000.copy()
per1000_score["每千人口熵权得分"] = X.mul(weight_map, axis=1).sum(axis=1)
per1000_score["每千人口排序"] = per1000_score["每千人口熵权得分"].rank(ascending=False, method="min").astype(int)
per1000_score = per1000_score.sort_values("每千人口熵权得分", ascending=False).reset_index(drop=True)

final_table = match_df.merge(per1000_score[["地区", "每千人口熵权得分", "每千人口排序"]], on="地区", how="left")
summary_counts = match_df["匹配类型"].value_counts().reindex(["严重失配", "相对失配", "基本匹配", "相对超前", "明显超前"]).fillna(0).astype(int).reset_index()
summary_counts.columns = ["匹配类型", "地区数"]

with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    aging_std.to_excel(writer, sheet_name="老龄化标准化矩阵", index=False)
    aging_equal.sort_values("U1排序").to_excel(writer, sheet_name="U1等权得分", index=False)
    med_all.to_excel(writer, sheet_name="医疗资源8指标标准化矩阵", index=False)
    u2_df.sort_values("U2排序").to_excel(writer, sheet_name="U2等权得分", index=False)
    per1000.to_excel(writer, sheet_name="每千人口5指标标准化矩阵", index=False)
    P.to_excel(writer, sheet_name="每千人口比例P", index=False)
    weights_df.to_excel(writer, sheet_name="每千人口熵权法权重", index=False)
    per1000_score.to_excel(writer, sheet_name="每千人口熵权得分排序", index=False)
    summary_counts.to_excel(writer, sheet_name="匹配类型统计", index=False)
    final_table.to_excel(writer, sheet_name="最终结果表", index=False)

print("已生成：", out_path)
