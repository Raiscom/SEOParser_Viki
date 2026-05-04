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


class XmlRiverDomainTopResult(BaseModel):
    """Описывает результат проверки домена в топе Yandex XMLRiver."""

    query: str
    position: str = ""
    url: str = ""
    domain: str = ""
    target_domain: str = ""
    error_message: str = ""
    error_code: str = ""


class DomainKeywordCandidate(BaseModel):
    """Describes a generated keyword candidate for a domain."""

    phrase: str
    source: str = ""
    score: float = 0.0
    source_url: str = ""
    title: str = ""
    h1: str = ""


class DomainKeywordPosition(BaseModel):
    """Describes a generated domain keyword and its Yandex position check."""

    phrase: str
    position: str = ""
    url: str = ""
    domain: str = ""
    source: str = ""
    score: float = 0.0
    source_url: str = ""
    title: str = ""
    h1: str = ""
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


class BukvarixKeywordResult(BaseModel):
    """Describes one Bukvarix keyword search row."""

    source_query: str
    keyword: str = ""
    words_count: str = ""
    chars_count: str = ""
    broad_frequency: str = ""
    exact_frequency: str = ""
    raw: str = ""
    error_message: str = ""
    error_code: str = ""


class BukvarixDomainResult(BaseModel):
    """Describes one Bukvarix domain search row."""

    source_domain: str
    keyword: str = ""
    words_count: str = ""
    chars_count: str = ""
    serp_results: str = ""
    broad_frequency: str = ""
    exact_frequency: str = ""
    position: str = ""
    raw: str = ""
    error_message: str = ""
    error_code: str = ""


class ProgressState(BaseModel):
    """Хранит состояние прогресса для интерфейса."""

    total: int = 0
    completed: int = 0
    failed: int = 0
    is_running: bool = False
    current_query: str = Field(default="")
