"""Настройка логирования приложения."""

from __future__ import annotations

from pathlib import Path

from loguru import logger

from app.config import get_app_dir


def setup_logging() -> None:
    """Настраивает логирование в файл рядом с приложением."""
    log_path = Path(get_app_dir(), "parser.log")
    logger.remove()
    logger.add(
        log_path,
        rotation="5 MB",
        retention="7 days",
        encoding="utf-8",
    )
