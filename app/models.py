"""Модели данных для результатов парсинга и состояния задач."""

from __future__ import annotations

from pydantic import BaseModel, Field


class XmlRiverResult(BaseModel):
    """Описывает одну строку результата XMLRiver."""

    query: str
    position: str = ""
    url: str = ""
    domain: str = ""
    title: str = ""
    snippet: str = ""
    error_message: str = ""
    error_code: str = ""


class SerpRiverResult(BaseModel):
    """Описывает одну строку результата SERPRiver."""

    query: str
    domain: str
    position: str = "Не найдено"
    url: str = ""
    title: str = ""
    error_message: str = ""
    error_code: str = ""
    raw_response: str = ""


class WordstatResult(BaseModel):
    """Описывает одну строку результата Wordstat."""

    query: str
    result_type: str = ""
    phrase: str = ""
    value: str = ""
    error_message: str = ""
    error_code: str = ""


class ProgressState(BaseModel):
    """Хранит состояние прогресса для интерфейса."""

    total: int = 0
    completed: int = 0
    failed: int = 0
    is_running: bool = False
    current_query: str = Field(default="")
