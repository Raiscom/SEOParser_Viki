"""Загрузка справочников для выпадающих списков."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from app.config import get_app_dir, get_bundle_dir


@dataclass(slots=True)
class ReferenceOption:
    """Описывает элемент выпадающего списка."""

    label: str
    value: str


def get_reference_dir() -> Path:
    """Возвращает каталог справочников рядом с .exe или внутри bundle."""
    app_reference_dir = get_app_dir() / "data" / "references"
    if app_reference_dir.exists():
        return app_reference_dir

    return get_bundle_dir() / "data" / "references"


REFERENCE_DIR = get_reference_dir()
REFERENCE_FILES = {
    "google_geo": REFERENCE_DIR / "geo.csv",
    "countries": REFERENCE_DIR / "countries.xlsx",
    "langs": REFERENCE_DIR / "langs.xlsx",
    "domains": REFERENCE_DIR / "domains.xlsx",
    "yandex_geo": REFERENCE_DIR / "yandex_geo.csv",
}


def load_reference_catalogs() -> tuple[dict[str, list[ReferenceOption]], dict[str, str]]:
    """Загружает справочники dropdown-ов и ошибки загрузки."""
    catalogs: dict[str, list[ReferenceOption]] = {}
    errors: dict[str, str] = {}
    loaders = {
        "google_geo": _load_google_geo,
        "countries": _load_countries,
        "langs": _load_langs,
        "domains": _load_domains,
        "yandex_geo": _load_yandex_geo,
    }
    for catalog_name, loader in loaders.items():
        try:
            catalogs[catalog_name] = loader(REFERENCE_FILES[catalog_name])
        except (FileNotFoundError, ValueError, OSError) as error:
            errors[catalog_name] = str(error)
    return catalogs, errors


def _load_google_geo(file_path: Path) -> list[ReferenceOption]:
    """Загружает список гео Google для параметра loc."""
    rows = _read_google_geo_rows(file_path)
    options: list[ReferenceOption] = []
    for row in rows:
        if row["CountryCode"] != "RU" or row["Status"] != "Active":
            continue
        label = (row["CanonicalName"] or row["Name"]).replace(",", ", ").strip()
        value = row["CriteriaId"]
        if not label or not value:
            continue
        options.append(ReferenceOption(label=label, value=value))
    if not options:
        raise ValueError("Справочник Google Geo пуст или не содержит активных локаций RU")
    return options


def _load_countries(file_path: Path) -> list[ReferenceOption]:
    """Загружает список стран Google для параметра country."""
    dataframe = _read_excel(file_path)
    return _to_options(dataframe, "place", "id")


def _load_langs(file_path: Path) -> list[ReferenceOption]:
    """Загружает список языков Google для параметра lr."""
    dataframe = _read_excel(file_path)
    return _to_options(dataframe, "name", "lang")


def _load_domains(file_path: Path) -> list[ReferenceOption]:
    """Загружает список доменов Google."""
    dataframe = _read_excel(file_path)
    return _to_options(dataframe, "domain", "id")


def _load_yandex_geo(file_path: Path) -> list[ReferenceOption]:
    """Загружает список регионов Яндекса."""
    dataframe = _read_csv(file_path)
    return _to_options(dataframe, "place", "id")


def _read_google_geo_rows(file_path: Path) -> list[dict[str, str]]:
    """Читает гибридный CSV Google Geo со строками в разных форматах."""
    if not file_path.exists():
        raise FileNotFoundError(f"Не найден справочник: {file_path}")
    field_names = (
        "CriteriaId",
        "Name",
        "CanonicalName",
        "ParentId",
        "CountryCode",
        "TargetType",
        "Status",
    )
    rows: list[dict[str, str]] = []
    for index, raw_line in enumerate(file_path.read_text(encoding="utf-8-sig").splitlines(), start=1):
        normalized_line = _normalize_google_geo_line(raw_line)
        if not normalized_line:
            continue
        parsed_row = next(csv.reader([normalized_line], delimiter=",", quotechar='"'))
        if len(parsed_row) != len(field_names):
            raise ValueError(f"Некорректный формат geo.csv в строке {index}: ожидалось 7 полей, получено {len(parsed_row)}")
        rows.append(
            {field_name: value.strip() for field_name, value in zip(field_names, parsed_row, strict=True)}
        )
    return rows


def _normalize_google_geo_line(raw_line: str) -> str:
    """Приводит строку Google Geo к виду, который понимает csv.reader."""
    line = raw_line.strip()
    if not line:
        return ""
    if line.startswith('"') and line.endswith('"'):
        return line[1:-1].replace('""', '"')
    return line


def _read_csv(file_path: Path) -> pd.DataFrame:
    """Читает CSV-файл справочника."""
    if not file_path.exists():
        raise FileNotFoundError(f"Не найден справочник: {file_path}")
    return pd.read_csv(file_path)


def _read_excel(file_path: Path) -> pd.DataFrame:
    """Читает XLSX-файл справочника."""
    if not file_path.exists():
        raise FileNotFoundError(f"Не найден справочник: {file_path}")
    return pd.read_excel(file_path, sheet_name=0)


def _to_options(dataframe: pd.DataFrame, label_column: str, value_column: str) -> list[ReferenceOption]:
    """Преобразует таблицу в список элементов dropdown-а."""
    if label_column not in dataframe.columns or value_column not in dataframe.columns:
        raise ValueError(
            f"В справочнике нет ожидаемых колонок: {label_column}, {value_column}",
        )
    options: list[ReferenceOption] = []
    for _, row in dataframe.iterrows():
        raw_label = str(row[label_column]).strip()
        raw_value = str(row[value_column]).strip()
        if not raw_label or not raw_value or raw_label.lower() == "nan" or raw_value.lower() == "nan":
            continue
        options.append(ReferenceOption(label=raw_label, value=raw_value))
    if not options:
        raise ValueError(f"Справочник пуст или не содержит валидных значений: {label_column}")
    return options
