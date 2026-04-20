"""Настройки приложения и подготовка путей для source/portable запуска."""

from __future__ import annotations

import os
import sys
from functools import lru_cache
from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


def get_app_dir() -> Path:
    """Возвращает каталог запуска приложения или .exe."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_bundle_dir() -> Path:
    """Возвращает каталог ресурсов bundled-приложения или корень проекта."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent.parent


def configure_runtime_environment() -> None:
    """Настраивает окружение Tcl/Tk для portable-сборки."""
    if not getattr(sys, "frozen", False):
        return

    bundle_dir = get_bundle_dir()
    tcl_candidates = (
        bundle_dir / "_tcl_data",
        bundle_dir / "tcl" / "tcl8.6",
    )
    tk_candidates = (
        bundle_dir / "_tk_data",
        bundle_dir / "tcl" / "tk8.6",
    )

    for tcl_dir in tcl_candidates:
        if tcl_dir.exists():
            os.environ.setdefault("TCL_LIBRARY", str(tcl_dir))
            break

    for tk_dir in tk_candidates:
        if tk_dir.exists():
            os.environ.setdefault("TK_LIBRARY", str(tk_dir))
            break


class AppSettings(BaseSettings):
    """Хранит ключи API и базовые параметры приложения."""

    model_config = SettingsConfigDict(
        env_file=get_app_dir() / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    xmlriver_user_id: str = Field(default="", alias="XMLRIVER_USER_ID")
    xmlriver_api_key: str = Field(default="", alias="XMLRIVER_API_KEY")
    serpriver_api_key: str = Field(default="", alias="SERPRIVER_API_KEY")
    google_last_location_label: str = Field(default="", alias="GOOGLE_LAST_LOCATION_LABEL")
    yandex_last_region_label: str = Field(default="", alias="YANDEX_LAST_REGION_LABEL")
    wordstat_last_region_label: str = Field(default="", alias="WORDSTAT_LAST_REGION_LABEL")

    request_connect_timeout: int = 5
    request_read_timeout: int = 30
    xmlriver_max_concurrency: int = 10
    serpriver_max_concurrency: int = 10
    import_row_limit: int = 10_000


@lru_cache(maxsize=1)
def get_settings() -> AppSettings:
    """Возвращает кешированный экземпляр настроек."""
    return AppSettings()


def save_settings(settings: AppSettings) -> None:
    """Сохраняет ключи API в .env рядом с приложением."""
    env_path = get_app_dir() / ".env"
    env_lines = [
        f"XMLRIVER_USER_ID={settings.xmlriver_user_id}",
        f"XMLRIVER_API_KEY={settings.xmlriver_api_key}",
        f"SERPRIVER_API_KEY={settings.serpriver_api_key}",
        f"GOOGLE_LAST_LOCATION_LABEL={settings.google_last_location_label}",
        f"YANDEX_LAST_REGION_LABEL={settings.yandex_last_region_label}",
        f"WORDSTAT_LAST_REGION_LABEL={settings.wordstat_last_region_label}",
        f"REQUEST_CONNECT_TIMEOUT={settings.request_connect_timeout}",
        f"REQUEST_READ_TIMEOUT={settings.request_read_timeout}",
        f"XMLRIVER_MAX_CONCURRENCY={settings.xmlriver_max_concurrency}",
        f"SERPRIVER_MAX_CONCURRENCY={settings.serpriver_max_concurrency}",
        f"IMPORT_ROW_LIMIT={settings.import_row_limit}",
    ]
    env_path.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
    get_settings.cache_clear()
