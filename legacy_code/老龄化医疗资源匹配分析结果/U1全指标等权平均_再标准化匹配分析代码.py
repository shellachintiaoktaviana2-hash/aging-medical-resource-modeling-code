import os
import pandas as pd

base_dir = os.path.dirname(os.path.abspath(__file__))
aging_path = os.path.join(base_dir, "老龄化熵权法结果(2).xlsx")
medical_path = os.path.join(base_dir, "2024医疗资源配置_熵权TOPSIS_系统聚类结果(2).xlsx")
output_path = os.path.join(base_dir, "U1全指标等权平均_再标准化匹配分析结果_代码运行输出.xlsx")

aging_raw = pd.read_excel(aging_path, sheet_name="原始数据")
infra_std = pd.read_excel(medical_path, sheet_name="基础设施_标准化")
human_std = pd.read_excel(medical_path, sheet_name="人力资源_标准化")
service_std = pd.read_excel(medical_path, sheet_name="医疗服务_标准化")

aging_cols = [c for c in aging_raw.columns if c not in ["地区", "年份"]]
aging_std = aging_raw.copy()
for c in aging_cols:
    aging_std[c] = (aging_raw[c] - aging_raw[c].min()) / (aging_raw[c].max() - aging_raw[c].min())
aging_std["U1_原始系统得分"] = aging_std[aging_cols].mean(axis=1)

medical_std = infra_std.merge(human_std, on="地区").merge(service_std, on="地区")
medical_cols = [c for c in medical_std.columns if c != "地区"]
medical_std["U2_原始系统得分"] = medical_std[medical_cols].mean(axis=1)

result = aging_std[["地区", "年份"] + aging_cols + ["U1_原始系统得分"]].merge(
    medical_std[["地区"] + medical_cols + ["U2_原始系统得分"]],
    on="地区",
    how="inner"
)

result["U1_标准化系统得分"] = (
    (result["U1_原始系统得分"] - result["U1_原始系统得分"].min()) /
    (result["U1_原始系统得分"].max() - result["U1_原始系统得分"].min())
)
result["U2_标准化系统得分"] = (
    (result["U2_原始系统得分"] - result["U2_原始系统得分"].min()) /
    (result["U2_原始系统得分"].max() - result["U2_原始系统得分"].min())
)
result["Mi"] = result["U2_标准化系统得分"] - result["U1_标准化系统得分"]

def classify(m):
    if m < -0.20:
        return "严重失配"
    elif m < -0.05:
        return "相对失配"
    elif m <= 0.05:
        return "基本匹配"
    elif m <= 0.20:
        return "相对超前"
    return "明显超前"

result["匹配类型"] = result["Mi"].apply(classify)
result = result.sort_values("Mi").reset_index(drop=True)
result["Mi排序"] = result.index + 1

summary = (
    result["匹配类型"]
    .value_counts()
    .reindex(["严重失配", "相对失配", "基本匹配", "相对超前", "明显超前"])
    .fillna(0)
    .astype(int)
    .reset_index()
)
summary.columns = ["匹配类型", "地区数量"]

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="最终结果表", index=False)
    summary.to_excel(writer, sheet_name="分类汇总", index=False)
    aging_std.to_excel(writer, sheet_name="U1计算明细", index=False)
    medical_std.to_excel(writer, sheet_name="U2计算明细", index=False)

print("处理完成，结果已保存到：", output_path)
