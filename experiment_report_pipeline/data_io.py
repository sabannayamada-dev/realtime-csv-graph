from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd


SUPPORTED_EXTENSIONS = {".csv", ".txt", ".xlsx", ".xls"}


def detect_delimiter(path: Path) -> str:
    sample = path.read_text(encoding="utf-8-sig", errors="replace")[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t;").delimiter
    except csv.Error:
        first_line = sample.splitlines()[0] if sample.splitlines() else ""
        return "\t" if "\t" in first_line else ","


def read_data(path: Path, header: int | None = 0) -> pd.DataFrame:
    path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(f"入力ファイルが見つかりません: {path}")
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"未対応の入力形式です: {path.suffix}")

    if path.suffix.lower() in {".xlsx", ".xls"}:
        frame = pd.read_excel(path, header=header)
    else:
        delimiter = detect_delimiter(path)
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "cp932", "utf-8"):
            try:
                frame = pd.read_csv(
                    path,
                    header=header,
                    sep=delimiter,
                    encoding=encoding,
                    engine="python",
                )
                break
            except UnicodeDecodeError as exc:
                last_error = exc
        else:
            raise ValueError(f"文字コードを判定できません: {path}") from last_error

    frame = frame.dropna(how="all").dropna(axis=1, how="all")
    if frame.empty:
        raise ValueError(f"入力データが空です: {path}")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def numeric_series(frame: pd.DataFrame, column: str) -> pd.Series:
    if column not in frame.columns:
        available = ", ".join(map(str, frame.columns))
        raise ValueError(f"列 '{column}' がありません。利用可能な列: {available}")
    values = pd.to_numeric(frame[column], errors="coerce")
    if values.notna().sum() == 0:
        raise ValueError(f"列 '{column}' に数値がありません")
    return values
