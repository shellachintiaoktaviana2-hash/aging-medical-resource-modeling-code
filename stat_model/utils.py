from __future__ import annotations

from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


PROVINCE_COLUMNS = ["地区", "省份", "省级地区"]


def ensure_dir(path: str | Path) -> Path:
    out = Path(path)
    out.mkdir(parents=True, exist_ok=True)
    return out


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_column(df: pd.DataFrame, keywords: Iterable[str], required: bool = True) -> str | None:
    keys = list(keywords)
    for col in df.columns:
        name = str(col).strip()
        if all(k in name for k in keys):
            return str(col)
    if required:
        raise ValueError(f"未找到包含关键词 {keys} 的列，当前列名为：{list(df.columns)}")
    return None


def province_column(df: pd.DataFrame) -> str:
    for col in PROVINCE_COLUMNS:
        if col in df.columns:
            return col
    for col in df.columns:
        if "地区" in str(col) or "省份" in str(col):
            return str(col)
    raise ValueError(f"未找到地区/省份列，当前列名为：{list(df.columns)}")


def normalize_province_column(df: pd.DataFrame, target: str = "地区") -> pd.DataFrame:
    df = clean_column_names(df)
    col = province_column(df)
    out = df.rename(columns={col: target}).copy()
    out[target] = out[target].astype(str).str.strip()
    out = out[out[target].notna() & (out[target] != "") & (out[target] != "nan")]
    return out


def read_table_auto_header(
    path: str | Path,
    sheet_name: str,
    required_columns: Iterable[str],
    max_scan_rows: int = 30,
) -> pd.DataFrame:
    """Read an Excel sheet whose real header may be below title/comment rows."""
    path = Path(path)
    required = set(required_columns)
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    for idx, row in preview.iterrows():
        values = {str(v).strip() for v in row.dropna().tolist()}
        if required.issubset(values):
            df = pd.read_excel(path, sheet_name=sheet_name, header=idx)
            df = clean_column_names(df)
            return df.dropna(how="all")
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = clean_column_names(df)
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{path.name}/{sheet_name} 未找到列 {sorted(missing)}，当前列名：{list(df.columns)}")
    return df.dropna(how="all")


def minmax_standardize(series: pd.Series, direction: str = "positive", fill_equal: float = 0.0) -> pd.Series:
    x = pd.to_numeric(series, errors="coerce").astype(float)
    min_v = x.min()
    max_v = x.max()
    if pd.isna(min_v) or pd.isna(max_v) or np.isclose(max_v, min_v):
        return pd.Series(fill_equal, index=series.index, dtype=float)
    if direction == "positive":
        return (x - min_v) / (max_v - min_v)
    if direction == "negative":
        return (max_v - x) / (max_v - min_v)
    raise ValueError("direction must be 'positive' or 'negative'")


def efficacy_standardize(series: pd.Series, direction: str = "positive") -> pd.Series:
    z = minmax_standardize(series, direction=direction, fill_equal=1.0)
    return 0.1 + 0.9 * z


def load_score_table(
    path: str | Path,
    sheet_candidates: Iterable[str],
    score_keywords: Iterable[str],
    province_name: str = "地区",
) -> pd.DataFrame:
    path = Path(path)
    xls = pd.ExcelFile(path)
    chosen = None
    for sheet in sheet_candidates:
        if sheet in xls.sheet_names:
            chosen = sheet
            break
    if chosen is None:
        chosen = xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=chosen)
    df = normalize_province_column(df, province_name)
    score_col = find_column(df, score_keywords)
    keep = [province_name, score_col]
    rank_col = find_column(df, ["排名"], required=False)
    if rank_col and rank_col not in keep:
        keep.append(rank_col)
    out = df[keep].copy()
    out = out.rename(columns={score_col: "综合得分"})
    out["综合得分"] = pd.to_numeric(out["综合得分"], errors="coerce")
    out = out.dropna(subset=[province_name, "综合得分"])
    return out
