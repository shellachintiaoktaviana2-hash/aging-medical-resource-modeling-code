from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def first_existing(root: Path, patterns: list[str]) -> Path | None:
    for pattern in patterns:
        found = list(root.glob(pattern))
        if found:
            return found[0]
    return None


def run(cmd: list[str]) -> None:
    print("运行：", " ".join(cmd))
    subprocess.run(cmd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the whole modeling workflow when standard input files are available.")
    parser.add_argument("--data-dir", default="data/raw")
    parser.add_argument("--output-dir", default="outputs")
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parents[1]
    data_dir = (project_root / args.data_dir).resolve()
    output_dir = (project_root / args.output_dir).resolve()

    aging_file = first_existing(data_dir, ["*老龄化拓展指标数据表*.xlsx", "*老龄化*指标*.xlsx"])
    medical_file = first_existing(data_dir, ["*医疗资源配置核心变量提取*2024*.xlsx", "*医疗资源*核心*.xlsx"])
    land_file = first_existing(data_dir, ["*土地相关指标*熵权法*.xlsx"])
    quality_file = first_existing(data_dir, ["*医疗资源配置_AHP权重与结果*质量侧重版*.xlsx"])

    if not aging_file:
        raise FileNotFoundError("未在 data/raw 中找到老龄化指标数据表。")
    if not medical_file:
        raise FileNotFoundError("未在 data/raw 中找到医疗资源核心变量数据表。")

    python = sys.executable
    run([python, str(project_root / "scripts" / "01_aging_models.py"), "--aging-file", str(aging_file), "--output-dir", str(output_dir / "aging")])

    med_cmd = [python, str(project_root / "scripts" / "02_medical_resource_models.py"), "--medical-file", str(medical_file), "--output-dir", str(output_dir / "medical")]
    if land_file:
        med_cmd += ["--land-file", str(land_file)]
    if quality_file:
        med_cmd += ["--quality-file", str(quality_file)]
    run(med_cmd)

    run([
        python,
        str(project_root / "scripts" / "03_match_and_plots.py"),
        "--aging-score",
        str(output_dir / "aging" / "aging_composite_index.xlsx"),
        "--medical-dir",
        str(output_dir / "medical"),
        "--output-dir",
        str(output_dir / "matching"),
    ])

    run([
        python,
        str(project_root / "scripts" / "04_robustness_checks.py"),
        "--match-dir",
        str(output_dir / "matching"),
        "--output-dir",
        str(output_dir / "robustness"),
    ])


if __name__ == "__main__":
    main()
