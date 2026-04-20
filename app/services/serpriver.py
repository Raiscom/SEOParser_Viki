"""Клиент SERPRiver для проверки позиций домена."""

from __future__ import annotations

import asyncio
from collections.abc import Callable
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree

import aiohttp
from aiohttp import ClientTimeout
from loguru import logger
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from app.models import SerpRiverResult

SERPRIVER_ENDPOINT = "https://serpriver.ru/api/search.php"


class SerpRiverClient:
    """Выполняет запросы к SERPRiver и ищет домен в выдаче."""

    def __init__(self, api_key: str, connect_timeout: int, read_timeout: int, max_concurrency: int) -> None:
        self.api_key = api_key
        self.timeout = ClientTimeout(connect=connect_timeout, sock_read=read_timeout)
        self.semaphore = asyncio.Semaphore(max_concurrency)

    async def fetch_queries(
        self,
        queries: list[str],
        target_domain: str,
        search_params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[SerpRiverResult]:
        """Проверяет список запросов и возвращает позиции домена."""
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            tasks = [
                asyncio.create_task(
                    self._fetch_indexed_query(session, index, query, target_domain, search_params),
                )
                for index, query in enumerate(queries)
            ]
            indexed_results: dict[int, SerpRiverResult] = {}
            total_queries = len(tasks)
            completed_queries = 0
            for task in asyncio.as_completed(tasks):
                index, result = await task
                indexed_results[index] = result
                completed_queries += 1
                if progress_callback is not None:
                    progress_callback(completed_queries, total_queries, result.query)
            return [indexed_results[index] for index in sorted(indexed_results)]

    async def _fetch_indexed_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        query: str,
        target_domain: str,
        search_params: dict[str, Any],
    ) -> tuple[int, SerpRiverResult]:
        """Возвращает индекс исходного запроса вместе с результатом."""
        return index, await self._fetch_single_query(session, query, target_domain, search_params)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(min=1, max=10),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
        reraise=True,
    )
    async def _fetch_single_query(
        self,
        session: aiohttp.ClientSession,
        query: str,
        target_domain: str,
        search_params: dict[str, Any],
    ) -> SerpRiverResult:
        """Отправляет один запрос и ищет домен в JSON- или XML-ответе."""
        engine = str(search_params.get("system", "yandex")).strip().lower()
        request_params = self._build_params(query, search_params)
        output_format = str(search_params.get("output_format", "json")).strip().lower() or "json"
        request_headers = {
            "accept": "application/xml" if output_format == "xml" else "application/json",
            "content-type": "application/json",
            "authorization": self.api_key,
        }
        async with self.semaphore:
            logger.info("Query sent | engine={} query={} page={}", engine, query, "")
            async with session.get(SERPRIVER_ENDPOINT, params=request_params, headers=request_headers) as response:
                if response.status != 200:
                    error_text = await response.text()
                    return SerpRiverResult(
                        query=query,
                        domain=target_domain,
                        error_code=str(response.status),
                        error_message=error_text[:300],
                        raw_response=error_text,
                    )
                response_text = await response.text()
                payload = self._deserialize_response(response_text)
        return self._extract_position(query, target_domain, payload, response_text)

    def _build_params(self, query: str, search_params: dict[str, Any]) -> dict[str, Any]:
        """Собирает query string для вызова SERPRiver."""
        engine = str(search_params.get("system", "yandex")).strip().lower()
        request_params: dict[str, Any] = {
            "api_key": self.api_key,
            "query": query,
            "system": engine,
            "domain": str(search_params.get("domain", "ru")).strip(),
            "result_cnt": str(search_params.get("result_cnt", "10")).strip(),
            "device": str(search_params.get("device", "desktop")).strip(),
            "output_format": str(search_params.get("output_format", "json")).strip().lower() or "json",
        }
        if engine == "google":
            request_params["location"] = str(search_params.get("location", "")).strip()
            request_params["hl"] = str(search_params.get("hl", "")).strip()
            request_params["gl"] = str(search_params.get("gl", "")).strip()
        else:
            request_params["lr"] = str(search_params.get("lr", "")).strip()
        return {key: value for key, value in request_params.items() if value}

    def _deserialize_response(self, response_text: str) -> Any:
        """Преобразует текст ответа в JSON-структуру или упрощённую XML-модель."""
        cleaned_text = response_text.strip()
        if not cleaned_text:
            return None
        if cleaned_text.startswith("<"):
            return self._parse_xml_response(cleaned_text)
        try:
            import json

            return json.loads(cleaned_text)
        except ValueError:
            logger.error("SERPRiver parse failed | error={} preview={}", "invalid_json", cleaned_text[:200])
            return {"code": "invalid_json", "message": cleaned_text[:300]}

    def _parse_xml_response(self, response_text: str) -> dict[str, Any]:
        """Преобразует XML-ответ в структуру, совместимую с общей обработкой."""
        try:
            root = ElementTree.fromstring(response_text)
        except ElementTree.ParseError:
            logger.error("SERPRiver parse failed | error={} preview={}", "invalid_xml", response_text[:200])
            return {"code": "invalid_xml", "message": response_text[:300]}

        error_payload = self._extract_xml_error(root)
        if error_payload is not None:
            return error_payload

        items: list[dict[str, Any]] = []
        for element in root.iter():
            item = self._build_item_from_xml_element(element)
            if item is not None:
                items.append(item)
        if items:
            return {"results": items}
        return {"code": "empty_xml", "message": "Пустой XML-ответ SERPRiver"}

    def _extract_xml_error(self, root: ElementTree.Element) -> dict[str, str] | None:
        """Извлекает код и текст ошибки из XML-ответа."""
        code_value = self._find_xml_text(root, {"code", "error_code"})
        if not code_value:
            return None
        message_value = self._find_xml_text(root, {"message", "error", "error_message"}) or "Ошибка API"
        return {"code": code_value, "message": message_value}

    def _build_item_from_xml_element(self, element: ElementTree.Element) -> dict[str, Any] | None:
        """Преобразует XML-узел результата в словарь позиции."""
        children = list(element)
        if not children:
            return None
        normalized_children = {child.tag.lower(): (child.text or "").strip() for child in children}
        has_link = bool(normalized_children.get("link") or normalized_children.get("url"))
        has_position = bool(normalized_children.get("position"))
        has_title = bool(normalized_children.get("title"))
        if not has_link and not has_position and not has_title:
            return None
        return {
            "position": normalized_children.get("position", ""),
            "link": normalized_children.get("link") or normalized_children.get("url", ""),
            "title": normalized_children.get("title", ""),
            "snippet": normalized_children.get("snippet", ""),
        }

    def _find_xml_text(self, root: ElementTree.Element, tag_names: set[str]) -> str:
        """Ищет первый непустой текст по набору tag-имен."""
        for element in root.iter():
            if element.tag.lower() not in tag_names:
                continue
            text_value = (element.text or "").strip()
            if text_value:
                return text_value
        return ""

    def _extract_position(self, query: str, target_domain: str, payload: Any, raw_response: str) -> SerpRiverResult:
        """Находит первое вхождение домена в списке результатов."""
        normalized_target = self._normalize_domain(target_domain)
        if payload is None:
            return SerpRiverResult(
                query=query,
                domain=target_domain,
                error_code="empty",
                error_message="Пустой ответ API",
                raw_response=raw_response,
            )
        if isinstance(payload, dict) and payload.get("code"):
            error_code = str(payload.get("code"))
            error_message = str(payload.get("message", "Ошибка API"))
            logger.error("API error | code={} query={} error={}", error_code, query, error_message)
            return SerpRiverResult(
                query=query,
                domain=target_domain,
                error_code=error_code,
                error_message=error_message,
                raw_response=raw_response,
            )
        if isinstance(payload, dict) and isinstance(payload.get("search_metadata"), dict):
            status = str(payload["search_metadata"].get("status", "")).strip().lower()
            if status and status != "success":
                error_message = str(payload["search_metadata"].get("message", "Ошибка API"))
                logger.error("API error | code={} query={} error={}", status, query, error_message)
                return SerpRiverResult(
                    query=query,
                    domain=target_domain,
                    error_code=status,
                    error_message=error_message,
                    raw_response=raw_response,
                )

        for item in self._extract_items(payload):
            url = str(item.get("link", item.get("url", ""))).strip()
            current_domain = self._normalize_domain(urlparse(url).netloc)
            if current_domain == normalized_target or current_domain.endswith(f".{normalized_target}"):
                return SerpRiverResult(
                    query=query,
                    domain=target_domain,
                    position=str(item.get("position", "")) or "Не найдено",
                    url=url,
                    title=str(item.get("title", "")).strip(),
                    raw_response=raw_response,
                )
        return SerpRiverResult(query=query, domain=target_domain, raw_response=raw_response)

    def _extract_items(self, payload: Any) -> list[dict[str, Any]]:
        """Достаёт список органических результатов из разных форм JSON."""
        if isinstance(payload, list):
            return [item for item in payload if isinstance(item, dict)]
        if not isinstance(payload, dict):
            return []
        for key in ("res", "results", "organic", "organic_results", "data"):
            value = payload.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
        return []

    def _normalize_domain(self, domain_value: str) -> str:
        """Нормализует домен для корректного сравнения."""
        cleaned_domain = domain_value.strip().lower()
        return cleaned_domain.removeprefix("www.")
