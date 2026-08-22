import numpy as np
import pandas as pd

INPUT_FILE = r"/mnt/data/医疗资源配置核心变量提取_2024_质量修订版.xlsx"
OUTPUT_NOTE = "本脚本展示质量侧重版AHP的权重计算，并把标准化方法从极差标准化修正为功效系数法。"

def ahp_weights(matrix):
    A = np.array(matrix, dtype=float)
    vals, vecs = np.linalg.eig(A)
    idx = np.argmax(vals.real)
    lam = float(vals[idx].real)
    w = np.abs(vecs[:, idx].real)
    w = w / w.sum()
    n = A.shape[0]
    RI = {1:0,2:0,3:0.58,4:0.90,5:1.12,6:1.24,7:1.32,8:1.41,9:1.45,10:1.49}
    ci = (lam - n) / (n - 1) if n > 1 else 0
    cr = ci / RI[n] if n > 2 and RI[n] != 0 else 0
    return w, lam, ci, cr

df = pd.read_excel(INPUT_FILE, sheet_name="Quality_Prov_2024", header=3)

criteria_matrix = np.array([
    [1,   1/3, 1/2, 2],
    [3,   1,   2,   4],
    [2,   1/2, 1,   3],
    [1/2, 1/4, 1/3, 1]
], dtype=float)
cw, _, _, _ = ahp_weights(criteria_matrix)

b1_matrix = np.array([
    [1,   1/3, 1/3],
    [3,   1,   1],
    [3,   1,   1]
], dtype=float)
b1w, _, _, _ = ahp_weights(b1_matrix)

b2_matrix = np.array([
    [1,   1/4, 1/3],
    [4,   1,   2],
    [3,   1/2, 1]
], dtype=float)
b2w, _, _, _ = ahp_weights(b2_matrix)

b3_matrix = np.array([
    [1,   1/3],
    [3,   1]
], dtype=float)
b3w, _, _, _ = ahp_weights(b3_matrix)

indicator_map = {
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

local_weight = {
    "C1_每千人口执业(助理)医师数": float(b1w[0]),
    "C2_每万人口全科医生数": float(b1w[1]),
    "C3_全科医生占执业(助理)医师比": float(b1w[2]),
    "C4_三级医院占医院比": float(b2w[0]),
    "C5_三级甲等医院占医院比": float(b2w[1]),
    "C6_三级甲等占三级医院比": float(b2w[2]),
    "C7_每十万人三级医院数": float(b3w[0]),
    "C8_每十万人三级甲等医院数": float(b3w[1]),
    "C9_信息化高配综合指数": 1.0,
}
parent_weight = {
    "C1_每千人口执业(助理)医师数": float(cw[0]),
    "C2_每万人口全科医生数": float(cw[0]),
    "C3_全科医生占执业(助理)医师比": float(cw[0]),
    "C4_三级医院占医院比": float(cw[1]),
    "C5_三级甲等医院占医院比": float(cw[1]),
    "C6_三级甲等占三级医院比": float(cw[1]),
    "C7_每十万人三级医院数": float(cw[2]),
    "C8_每十万人三级甲等医院数": float(cw[2]),
    "C9_信息化高配综合指数": float(cw[3]),
}
global_weight = {k: parent_weight[k] * local_weight[k] for k in indicator_map}
s = sum(global_weight.values())
global_weight = {k: v/s for k, v in global_weight.items()}

def efficacy_standardize(series):
    x = pd.to_numeric(series, errors="coerce")
    return 0.1 + 0.9 * (x - x.min()) / (x.max() - x.min())

score_df = df[["地区"]].copy()
for ind, col in indicator_map.items():
    score_df[ind] = efficacy_standardize(df[col])

score_df["综合得分"] = sum(score_df[k] * global_weight[k] for k in global_weight)
score_df["排名"] = score_df["综合得分"].rank(ascending=False, method="min").astype(int)
score_df = score_df.sort_values(["排名", "综合得分"], ascending=[True, False])

gd = score_df.loc[score_df["地区"] == "广东"].iloc[0]
raw_c1 = float(df.loc[df["地区"] == "广东", "每千人口执业(助理)医师数_计算"].iloc[0])

print("广东原始C1 =", raw_c1)
print("广东修正后标准化C1 =", round(float(gd["C1_每千人口执业(助理)医师数"]), 4))
print("\n前10位省份：")
print(score_df[["地区", "综合得分", "排名"]].head(10).to_string(index=False))
