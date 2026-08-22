from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import LinearSegmentedColormap
from scipy.cluster.hierarchy import dendrogram, linkage

from .utils import ensure_dir


def setup_chinese_font() -> None:
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Arial Unicode MS", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


def plot_gap_bar(df: pd.DataFrame, output_file: str | Path, title: str) -> None:
    setup_chinese_font()
    temp = df.sort_values("匹配差值", ascending=True)
    colors = np.where(temp["匹配差值"] < 0, "#D95F59", "#4E9F70")
    fig, ax = plt.subplots(figsize=(13, 7))
    ax.bar(temp["地区"], temp["匹配差值"], color=colors, alpha=0.88)
    ax.axhline(0, color="#333333", linewidth=1)
    ax.set_title(title, fontsize=16, fontweight="bold")
    ax.set_ylabel("匹配差值 U2-U1")
    ax.tick_params(axis="x", rotation=60)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    fig.savefig(output_file, dpi=220, bbox_inches="tight")
    plt.close(fig)


def plot_quadrant(df: pd.DataFrame, output_file: str | Path, title: str) -> None:
    setup_chinese_font()
    x_line = df["老龄化指数"].mean()
    y_line = df["医疗资源指数"].mean()
    color_map = {
        "高老龄-高资源": "#4E79A7",
        "高老龄-低资源": "#E15759",
        "低老龄-低资源": "#76B7B2",
        "低老龄-高资源": "#F28E2B",
    }
    fig, ax = plt.subplots(figsize=(10, 7.2))
    for label, group in df.groupby("四象限类型"):
        ax.scatter(group["老龄化指数"], group["医疗资源指数"], s=80, label=label, color=color_map.get(label, "#777777"), alpha=0.88)
    for _, row in df.iterrows():
        ax.text(row["老龄化指数"] + 0.006, row["医疗资源指数"] + 0.006, str(row["地区"]), fontsize=8)
    ax.axvline(x_line, color="#555555", linestyle="--", linewidth=1)
    ax.axhline(y_line, color="#555555", linestyle="--", linewidth=1)
    ax.set_xlabel("老龄化指数 U1")
    ax.set_ylabel("医疗资源配置指数 U2")
    ax.set_title(title, fontsize=16, fontweight="bold")
    ax.grid(alpha=0.25)
    ax.legend(frameon=False, fontsize=9)
    fig.tight_layout()
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    fig.savefig(output_file, dpi=220, bbox_inches="tight")
    plt.close(fig)


def plot_gap_heatmap(compare_df: pd.DataFrame, output_file: str | Path) -> None:
    setup_chinese_font()
    gap_cols = [c for c in compare_df.columns if c.endswith("_匹配差值")]
    data = compare_df[["地区"] + gap_cols].copy()
    data["平均匹配差值"] = data[gap_cols].mean(axis=1)
    data = data.sort_values("平均匹配差值").drop(columns=["平均匹配差值"])
    matrix = data.set_index("地区")
    values = matrix.to_numpy(dtype=float)
    max_abs = np.nanmax(np.abs(values))
    cmap = LinearSegmentedColormap.from_list("gap", ["#C0392B", "#F7F7F7", "#1E8449"])
    fig, ax = plt.subplots(figsize=(9, max(7, 0.34 * len(matrix))))
    im = ax.imshow(values, aspect="auto", cmap=cmap, vmin=-max_abs, vmax=max_abs)
    ax.set_xticks(np.arange(len(matrix.columns)))
    ax.set_xticklabels([c.replace("_匹配差值", "") for c in matrix.columns], fontsize=10)
    ax.set_yticks(np.arange(len(matrix.index)))
    ax.set_yticklabels(matrix.index, fontsize=8)
    for i in range(values.shape[0]):
        for j in range(values.shape[1]):
            ax.text(j, i, f"{values[i, j]:.2f}", ha="center", va="center", fontsize=7)
    ax.set_title("三种医疗资源口径下的匹配差值热力图", fontsize=15, fontweight="bold")
    fig.colorbar(im, ax=ax, fraction=0.035, pad=0.02)
    fig.tight_layout()
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    fig.savefig(output_file, dpi=220, bbox_inches="tight")
    plt.close(fig)


def plot_cluster_dendrogram(std_df: pd.DataFrame, indicators: list[str], output_file: str | Path) -> None:
    setup_chinese_font()
    labels = std_df["地区"].astype(str).tolist()
    X = std_df[indicators].astype(float).to_numpy()
    Z = linkage(X, method="ward")
    fig, ax = plt.subplots(figsize=(13, 7))
    dendrogram(Z, labels=labels, leaf_rotation=60, leaf_font_size=9, ax=ax)
    ax.set_title("基于人口口径五指标的医疗资源系统聚类结果", fontsize=15, fontweight="bold")
    ax.set_ylabel("Ward 距离")
    fig.tight_layout()
    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    fig.savefig(output_file, dpi=220, bbox_inches="tight")
    plt.close(fig)
