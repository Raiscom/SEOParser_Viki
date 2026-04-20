"""Клиент XMLRiver Wordstat New."""

from __future__ import annotations

import asyncio
from collections.abc import Callable
from typing import Any

import aiohttp
from aiohttp import ClientTimeout
from loguru import logger
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from app.models import WordstatResult

WORDSTAT_ENDPOINT = "https://xmlriver.com/wordstat/new/json"


class WordstatClient:
    """Выполняет запросы к Wordstat New и нормализует ответ."""

    def __init__(self, user_id: str, api_key: str, connect_timeout: int, read_timeout: int, max_concurrency: int) -> None:
        self.user_id = user_id
        self.api_key = api_key
        self.timeout = ClientTimeout(connect=connect_timeout, sock_read=read_timeout)
        self.semaphore = asyncio.Semaphore(max_concurrency)

    async def fetch_queries(
        self,
        queries: list[str],
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[WordstatResult]:
        """Выполняет пакет Wordstat-запросов и возвращает плоский список результатов."""
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            tasks = [
                asyncio.create_task(self._fetch_indexed_query(session, index, query, params))
                for index, query in enumerate(queries)
            ]
            indexed_results: dict[int, list[WordstatResult]] = {}
            total_queries = len(tasks)
            completed_queries = 0
            for task in asyncio.as_completed(tasks):
                index, query_results = await task
                indexed_results[index] = query_results
                completed_queries += 1
                current_query = query_results[0].query if query_results else ""
                if progress_callback is not None:
                    progress_callback(completed_queries, total_queries, current_query)
        ordered_results = [indexed_results[index] for index in sorted(indexed_results)]
        return [item for group in ordered_results for item in group]

    async def _fetch_indexed_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        query: str,
        params: dict[str, Any],
    ) -> tuple[int, list[WordstatResult]]:
        """Возвращает индекс исходного запроса вместе с его результатами."""
        return index, await self._fetch_single_query(session, query, params)

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
        params: dict[str, Any],
    ) -> list[WordstatResult]:
        """Отправляет один запрос Wordstat и нормализует JSON-ответ."""
        request_params = self._build_params(query, params)
        async with self.semaphore:
            logger.info("Query sent | engine={} query={} page={}", "wordstat", query, "")
            async with session.get(WORDSTAT_ENDPOINT, params=request_params) as response:
                if response.status != 200:
                    error_text = await response.text()
                    return [WordstatResult(query=query, error_code=str(response.status), error_message=error_text[:300])]
                payload = await response.json(content_type=None)
        return self._parse_wordstat_response(query, payload)

    def _build_params(self, query: str, params: dict[str, Any]) -> dict[str, Any]:
        """Формирует GET-параметры Wordstat."""
        request_params = {
            "user": self.user_id,
            "key": self.api_key,
            "query": query,
            "regions": params.get("regions", ""),
            "device": params.get("device", ""),
            "period": params.get("period", ""),
            "start": params.get("start", ""),
            "end": params.get("end", ""),
            "pagetype": params.get("pagetype", "words"),
        }
        return {key: value for key, value in request_params.items() if value or key == "device"}

    def _parse_wordstat_response(self, query: str, payload: Any) -> list[WordstatResult]:
        """Преобразует Wordstat JSON в список строк таблицы."""
        if not isinstance(payload, dict):
            return [WordstatResult(query=query, error_code="invalid_json", error_message="Некорректный JSON-ответ")]
        if "graph" in payload:
            return self._parse_history_response(query, payload)

        result_rows: list[WordstatResult] = []
        for item in payload.get("popular", []):
            result_rows.append(self._build_result_row(query, item, "popular"))
        for item in payload.get("associations", []):
            result_rows.append(self._build_result_row(query, item, "associations"))
        if result_rows:
            return result_rows
        return [WordstatResult(query=query, error_code="empty", error_message="Пустой ответ Wordstat")]

    def _parse_history_response(self, query: str, payload: dict[str, Any]) -> list[WordstatResult]:
        """Нормализует ответ Wordstat с pagetype=history."""
        table_data = payload.get("graph", {}).get("tableData", [])
        if not isinstance(table_data, list):
            return [WordstatResult(query=query, error_code="history_format", error_message="Некорректный формат history-ответа")]

        result_rows: list[WordstatResult] = []
        total_value = payload.get("totalValue")
        if total_value is not None:
            result_rows.append(
                WordstatResult(query=query, result_type="total", phrase="Общее значение", value=str(total_value)),
            )

        for item in table_data:
            if not isinstance(item, dict):
                continue
            absolute_value = str(item.get("absoluteValue", "")).strip()
            relative_value = str(item.get("value", "")).strip()
            combined_value = absolute_value
            if relative_value:
                combined_value = f"{absolute_value} | {relative_value}" if absolute_value else relative_value
            result_rows.append(
                WordstatResult(
                    query=query,
                    result_type="history",
                    phrase=str(item.get("text", "")).strip(),
                    value=combined_value,
                ),
            )

        if result_rows:
            return result_rows
        return [WordstatResult(query=query, error_code="empty", error_message="Пустой ответ Wordstat")]

    def _build_result_row(self, query: str, item: Any, result_type: str) -> WordstatResult:
        """Строит строку таблицы Wordstat."""
        if not isinstance(item, dict):
            return WordstatResult(
                query=query,
                result_type=result_type,
                error_code="invalid_item",
                error_message="Некорректный элемент ответа",
            )
        return WordstatResult(
            query=query,
            result_type=result_type,
            phrase=str(item.get("text", "")).strip(),
            value=str(item.get("value", "")).strip(),
        )
