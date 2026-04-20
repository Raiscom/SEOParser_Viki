"""Импорт запросов и доменов из CSV/XLSX."""

from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd


def read_csv_lines(file_path: Path, row_limit: int) -> list[str]:
    """Читает первый столбец CSV с автодетектом разделителя."""
    sample = file_path.read_text(encoding="utf-8-sig")
    dialect = csv.Sniffer().sniff(sample[:2048], delimiters=",;")
    rows: list[str] = []
    with file_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.reader(csv_file, delimiter=dialect.delimiter)
        for row in reader:
            if not row:
                continue
            value = row[0].strip()
            if value:
                rows.append(value)
    _validate_row_limit(rows, row_limit)
    return rows


def read_xlsx_lines(file_path: Path, row_limit: int) -> list[str]:
    """Читает первый лист и первый столбец XLSX."""
    dataframe = pd.read_excel(file_path, sheet_name=0, header=None, usecols=[0], engine="openpyxl")
    rows = [str(value).strip() for value in dataframe.iloc[:, 0].tolist() if str(value).strip()]
    _validate_row_limit(rows, row_limit)
    return rows


def _validate_row_limit(rows: list[str], row_limit: int) -> None:
    """Проверяет ограничение на количество импортируемых строк."""
    if len(rows) > row_limit:
        raise ValueError(f"Превышен лимит импорта: {len(rows)} > {row_limit}")
