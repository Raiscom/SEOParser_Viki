"""Экспорт результатов в CSV/XLSX."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


def build_export_path(directory: Path, prefix: str, suffix: str) -> Path:
    """Строит путь вида export_name_YYYYMMDD_HHMMSS.ext."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return directory / f"{prefix}_{timestamp}.{suffix}"


def export_to_csv(rows: list[dict[str, Any]], file_path: Path) -> None:
    """Сохраняет строки в CSV с BOM для Excel."""
    dataframe = pd.DataFrame(rows)
    dataframe.to_csv(file_path, index=False, encoding="utf-8-sig")


def export_to_xlsx(rows: list[dict[str, Any]], file_path: Path) -> None:
    """Сохраняет строки в XLSX."""
    dataframe = pd.DataFrame(rows)
    dataframe.to_excel(file_path, index=False, engine="openpyxl")
