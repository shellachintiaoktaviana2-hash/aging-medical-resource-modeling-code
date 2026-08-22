# Aging and Medical Resource Matching Models

本目录整理自统计建模论文《基于熵权法与 AHP 综合评价的省级地区老龄化与医疗资源配置匹配耦合研究》及过往对话中保留下来的脚本。

## 目录说明

- `stat_model/`: 整理后的公共函数，包括熵权法、AHP、匹配分析、稳健性检验和绘图。
- `scripts/`: 可直接运行的命令行脚本。
- `legacy_code/`: 从过往对话和附件包中恢复的原始脚本备份，未改动逻辑，用于溯源。
- `data/raw/`: 放置原始或中间 Excel 数据文件。
- `outputs/`: 默认输出结果目录。

## 对应论文模型

| 论文内容 | 整理后脚本 |
| --- | --- |
| 老龄化熵权法综合压力模型 | `scripts/01_aging_models.py` |
| AHP 老龄化质量模型 | `scripts/01_aging_models.py` |
| 老龄化熵权法与 AHP 合成 U1 | `scripts/01_aging_models.py` |
| 人口口径五指标熵权法 | `scripts/02_medical_resource_models.py` |
| 土地口径五指标熵权法 | `scripts/02_medical_resource_models.py` |
| 质量口径 AHP 模型 | `scripts/02_medical_resource_models.py` |
| U1 与三类 U2 匹配差值、四象限 | `scripts/03_match_and_plots.py` |
| 三口径稳健性、相关性、重合率 | `scripts/04_robustness_checks.py` |

## 安装依赖

```bash
pip install -r requirements.txt
```

## 推荐运行顺序

将数据文件放入 `data/raw/` 后运行：

```bash
python scripts/01_aging_models.py --aging-file data/raw/老龄化拓展指标数据表.xlsx --output-dir outputs/aging
python scripts/02_medical_resource_models.py --medical-file data/raw/医疗资源配置核心变量提取_2024.xlsx --output-dir outputs/medical
python scripts/03_match_and_plots.py --aging-score outputs/aging/aging_composite_index.xlsx --medical-dir outputs/medical --output-dir outputs/matching
python scripts/04_robustness_checks.py --match-dir outputs/matching --output-dir outputs/robustness
```

如果暂时没有完整原始数据，但已有论文附件中的评分结果 Excel，也可以直接用 `03_match_and_plots.py` 输入已有的老龄化最终得分、每千人口、土地口径、质量口径结果表继续做匹配分析。

## 旧脚本说明

`legacy_code/` 中保留了旧版代码，文件名和原始目录结构尽量不变。旧脚本中有些路径写死为当时电脑路径，整理版已经改为命令行参数和自动识别列名，后续上传 GitHub 时建议以 `scripts/` 和 `stat_model/` 为主。

