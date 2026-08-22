import pandas as pd
import numpy as np
from pathlib import Path

# =========================
# 1. 文件路径与工作表名
# =========================
FILE_PATH = r"D:\文献\老龄化指标提取_中文最终版_扩展指标版.xlsx"
SHEET_NAME = "老龄化扩展候选指标_2023"
OUTPUT_DIR = Path(FILE_PATH).parent

# =========================
# 2. 拓展模型指标设置
#    说明：
#    positive = 值越大，老龄化/老龄化压力越强
#    negative = 值越大，人口越年轻或再生产能力越强
# =========================
INDICATORS = {
    "0-14岁人口占比": "negative",
    "15-64岁人口占比": "negative",
    "65岁及以上人口占比": "positive",
    "老年抚养比": "positive",
    "总抚养比": "positive",
    "老少比": "positive",
    "出生率_‰": "negative",
    "自然增长率_‰": "negative",
}


# =========================
# 3. 读取并清洗数据
# =========================
def load_and_clean_data(file_path: str, sheet_name: str) -> pd.DataFrame:
    # 这张表前面有说明文字，因此 header=2 先读入，再删掉重复表头行
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)

    # 统一列名：真正列名在第一行数据里
    df.columns = df.iloc[0]
    df = df.iloc[1:].copy()

    # 删除空行
    df = df.dropna(how="all")

    # 去掉可能残留的重复表头行
    df = df[df["地区"].astype(str) != "地区"].copy()

    # 重置索引
    df.reset_index(drop=True, inplace=True)

    # 数值列转为数值型
    for col in df.columns:
        if col not in ["地区", "年份"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 年份转为整数（若需要）
    df["年份"] = pd.to_numeric(df["年份"], errors="coerce")

    # 删除关键列缺失的样本
    required_cols = ["地区", "年份"] + list(INDICATORS.keys())
    df = df.dropna(subset=required_cols).copy()

    return df


# =========================
# 4. 极差标准化
# =========================
def min_max_standardize(df: pd.DataFrame, indicator_info: dict, eps: float = 1e-12) -> pd.DataFrame:
    result = pd.DataFrame(index=df.index)

    for col, direction in indicator_info.items():
        x = df[col].astype(float)
        x_min = x.min()
        x_max = x.max()

        # 防止某指标所有地区完全相同导致分母为0
        if np.isclose(x_max, x_min):
            result[col] = 1.0
            continue

        if direction == "positive":
            z = (x - x_min) / (x_max - x_min)
        elif direction == "negative":
            z = (x_max - x) / (x_max - x_min)
        else:
            raise ValueError(f"指标 {col} 的方向只能是 'positive' 或 'negative'")

        # 避免后续 ln(0)
        result[col] = z + eps

    return result


# =========================
# 5. 熵权法计算
# =========================
def entropy_weight_method(std_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series, pd.Series]:
    n, m = std_df.shape

    # 比重矩阵
    p = std_df.div(std_df.sum(axis=0), axis=1)

    # 熵值
    k = 1 / np.log(n)
    e = -k * (p * np.log(p)).sum(axis=0)

    # 差异系数
    d = 1 - e

    # 权重
    w = d / d.sum()

    # 综合得分
    score = std_df.mul(w, axis=1).sum(axis=1)

    return p, e, w, score


# =========================
# 6. 主程序
# =========================
def main():
    print("开始读取并处理数据...")
    df = load_and_clean_data(FILE_PATH, SHEET_NAME)

    print("\n成功读取数据，前5行如下：")
    print(df[["地区", "年份"] + list(INDICATORS.keys())].head())

    # 标准化
    std_df = min_max_standardize(df, INDICATORS)

    # 熵权法
    p, e, w, score = entropy_weight_method(std_df)

    # 结果表
    result = df[["地区", "年份"]].copy()
    result["老龄化拓展模型得分"] = score
    result = result.sort_values(by="老龄化拓展模型得分", ascending=False).reset_index(drop=True)
    result["排名"] = result.index + 1
    result = result[["排名", "地区", "年份", "老龄化拓展模型得分"]]

    # 权重表
    weight_df = pd.DataFrame({
        "指标": w.index,
        "权重": w.values,
        "指标方向": [INDICATORS[i] for i in w.index],
        "熵值": e.values,
        "差异系数": (1 - e).values,
    }).sort_values(by="权重", ascending=False).reset_index(drop=True)

    # 标准化后的数据表
    std_output = df[["地区", "年份"]].join(std_df)

    # 输出到 Excel
    output_file = OUTPUT_DIR / "老龄化拓展模型_熵权法结果.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="综合得分排名", index=False)
        weight_df.to_excel(writer, sheet_name="指标权重", index=False)
        df.to_excel(writer, sheet_name="清洗后原始数据", index=False)
        std_output.to_excel(writer, sheet_name="标准化数据", index=False)
        p.to_excel(writer, sheet_name="比重矩阵", index=True)

    # 同时导出 CSV，便于直接查看
    result.to_csv(OUTPUT_DIR / "老龄化拓展模型_综合得分排名.csv", index=False, encoding="utf-8-sig")
    weight_df.to_csv(OUTPUT_DIR / "老龄化拓展模型_指标权重.csv", index=False, encoding="utf-8-sig")

    print("\n========== 熵权法计算完成 ==========")
    print("\n【指标权重】")
    print(weight_df)
    print("\n【综合得分前10名】")
    print(result.head(10))
    print(f"\n结果已保存到：{output_file}")


if __name__ == "__main__":
    main()
