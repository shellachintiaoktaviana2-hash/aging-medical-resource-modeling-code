import argparse
import pandas as pd

def min_max_standardize(df, cols):
    out = df[['地区']].copy()
    for col in cols:
        out[col] = (df[col] - df[col].min()) / (df[col].max() - df[col].min())
    return out

def classify_match(mi):
    if mi < -0.20:
        return '严重失配'
    elif mi < -0.05:
        return '相对失配'
    elif mi <= 0.05:
        return '基本匹配'
    elif mi <= 0.20:
        return '相对超前'
    else:
        return '明显超前'

def main(aging_file, med_file, output_file):
    aging_raw = pd.read_excel(aging_file, sheet_name='原始数据')
    aging_cols = ['65岁及以上人口占比', '老年抚养比', '老少比']
    aging_std = min_max_standardize(aging_raw, aging_cols)
    aging_std['U1_老龄化指数'] = aging_std[aging_cols].mean(axis=1)

    base_std = pd.read_excel(med_file, sheet_name='基础设施_标准化')
    hr_std = pd.read_excel(med_file, sheet_name='人力资源_标准化')
    serv_std = pd.read_excel(med_file, sheet_name='医疗服务_标准化')

    med_std = base_std.merge(hr_std, on='地区').merge(serv_std, on='地区')
    med_cols = [c for c in med_std.columns if c != '地区']
    med_std['U2_医疗资源配置指数'] = med_std[med_cols].mean(axis=1)

    result = aging_std[['地区', 'U1_老龄化指数']].merge(
        med_std[['地区', 'U2_医疗资源配置指数']], on='地区', how='inner'
    )
    result['U1排名'] = result['U1_老龄化指数'].rank(ascending=False, method='min').astype(int)
    result['U2排名'] = result['U2_医疗资源配置指数'].rank(ascending=False, method='min').astype(int)
    result['Mi_匹配差值'] = result['U2_医疗资源配置指数'] - result['U1_老龄化指数']
    result['匹配类型'] = result['Mi_匹配差值'].apply(classify_match)
    result = result[['地区', 'U1_老龄化指数', 'U1排名', 'U2_医疗资源配置指数', 'U2排名', 'Mi_匹配差值', '匹配类型']]
    result = result.sort_values('Mi_匹配差值').reset_index(drop=True)

    summary = result['匹配类型'].value_counts().reindex(
        ['严重失配', '相对失配', '基本匹配', '相对超前', '明显超前']
    ).fillna(0).astype(int).reset_index()
    summary.columns = ['匹配类型', '地区数']

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        aging_std.to_excel(writer, index=False, sheet_name='老龄化标准化')
        med_std.to_excel(writer, index=False, sheet_name='医疗资源标准化')
        result.to_excel(writer, index=False, sheet_name='最终结果表')
        summary.to_excel(writer, index=False, sheet_name='类型汇总')

    print('处理完成，输出文件：', output_file)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--aging_file', default='/mnt/data/老龄化熵权法结果(2).xlsx')
    parser.add_argument('--med_file', default='/mnt/data/2024医疗资源配置_熵权TOPSIS_系统聚类结果(2).xlsx')
    parser.add_argument('--output_file', default='/mnt/data/老龄化_医疗资源匹配差值分析结果.xlsx')
    args = parser.parse_args()
    main(args.aging_file, args.med_file, args.output_file)
