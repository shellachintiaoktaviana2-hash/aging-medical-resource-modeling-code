# Code Inventory

## Cleaned Code

| File | Purpose |
| --- | --- |
| `stat_model/utils.py` | Excel 自动表头识别、列名定位、地区列统一、极差标准化等工具函数 |
| `stat_model/entropy.py` | 熵权法计算，支持普通 0-1 标准化和土地口径使用的 0.1 下限功效系数法 |
| `stat_model/ahp.py` | 老龄化 AHP、医疗质量 AHP、判断矩阵权重和一致性检验 |
| `stat_model/matching.py` | 匹配差值、匹配类型、四象限、耦合度、三口径稳健性检验 |
| `stat_model/plotting.py` | 匹配差值条形图、四象限图、热力图、系统聚类图 |
| `scripts/01_aging_models.py` | 运行老龄化熵权法、老龄化 AHP，并合成最终 U1 指数 |
| `scripts/02_medical_resource_models.py` | 运行人口口径、土地口径、质量口径医疗资源评价 |
| `scripts/03_match_and_plots.py` | 运行 U1 与三类 U2 的匹配分析并输出图表 |
| `scripts/04_robustness_checks.py` | 输出三口径稳健性、排序相关和重点地区重合率 |
| `scripts/run_all.py` | 当标准数据文件都放在 `data/raw/` 时，一键运行完整流程 |

## Legacy Code Restored From Earlier Work

`legacy_code/` 保留 15 个旧版 `.py` 脚本，主要来自论文建模过程中的不同阶段：

- 老龄化熵权法与 AHP 建模脚本
- 老龄化最终综合得分合成脚本
- 每千人口医疗资源熵权法脚本
- 土地口径医疗资源熵权法脚本
- 医疗质量 AHP 权重计算脚本
- 老龄化与医疗资源人口、土地、质量口径匹配分析脚本
- 模型检验、稳健性和作图脚本

旧脚本已完整备份，但存在硬编码路径；正式复现与 GitHub 展示建议使用整理后的 `scripts/` 和 `stat_model/`。

## Verification

已用论文附件中的 Excel 数据跑通整理版流程，关键输出与论文描述一致：

- 老龄化综合指数前 5 位：辽宁、山东、重庆、四川、江苏。
- 人口口径匹配：严重失配 13 个，高老龄-低资源 7 个。
- 土地口径匹配：严重失配 22 个，高老龄-低资源 11 个。
- 质量口径匹配：严重失配 17 个，高老龄-低资源 11 个。

