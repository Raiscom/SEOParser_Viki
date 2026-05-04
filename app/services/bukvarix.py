"""Bukvarix API client."""

from __future__ import annotations

import asyncio
import csv
import json
from collections.abc import Callable, Iterable
from io import StringIO
from typing import Any

import aiohttp
from aiohttp import ClientTimeout
from loguru import logger
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from app.models import BukvarixDomainResult, BukvarixKeywordResult

BUKVARIX_KEYWORDS_ENDPOINT = "http://api.bukvarix.com/v1/keywords/"
BUKVARIX_MKEYWORDS_ENDPOINT = "http://api.bukvarix.com/v1/mkeywords/"
BUKVARIX_SITE_ENDPOINT = "http://api.bukvarix.com/v1/site/"
BUKVARIX_SITE_CMP_ENDPOINT = "http://api.bukvarix.com/v1/site_cmp/"
BUKVARIX_SITE_MCMP_ENDPOINT = "http://api.bukvarix.com/v1/site_mcmp/"

COMMON_OPTION_NAMES = ("num", "format", "bom", "header", "json_type", "result_count")
KEYWORD_OPTION_NAMES = (
    "q2",
    "report_type",
    "broad_from",
    "broad_to",
    "exact_from",
    "exact_to",
    "length_from",
    "length_to",
    "words_from",
    "words_to",
)
DOMAIN_OPTION_NAMES = ("q2", "region", "comparison_type")


class BukvarixClient:
    """Runs Bukvarix keyword and domain API requests."""

    def __init__(self, api_key: str, connect_timeout: int, read_timeout: int, max_concurrency: int) -> None:
        self.api_key = api_key
        self.timeout = ClientTimeout(connect=connect_timeout, sock_read=read_timeout)
        self.semaphore = asyncio.Semaphore(max_concurrency)

    async def fetch_keywords(
        self,
        queries: list[str],
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[BukvarixKeywordResult]:
        """Fetches keywords by one phrase or by the full phrase list."""
        mode = str(params.get("mode", "single")).strip().lower()
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            if mode == "multiple":
                results = await self._fetch_keyword_batch(session, queries, params)
                if progress_callback is not None:
                    progress_callback(1, 1, f"{len(queries)} phrases")
                return results

            tasks = [
                asyncio.create_task(self._fetch_indexed_keyword_query(session, index, query, params))
                for index, query in enumerate(queries)
            ]
            return await self._collect_indexed_results(tasks, progress_callback)

    async def fetch_domains(
        self,
        domains: list[str],
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[BukvarixDomainResult]:
        """Fetches keyword rows by domain or domain comparison mode."""
        mode = str(params.get("mode", "single")).strip().lower()
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            if mode == "multiple":
                results = await self._fetch_domain_batch(session, domains, params)
                if progress_callback is not None:
                    progress_callback(1, 1, f"{len(domains)} domains")
                return results

            tasks = [
                asyncio.create_task(self._fetch_indexed_domain_query(session, index, domain, params))
                for index, domain in enumerate(domains)
            ]
            return await self._collect_indexed_results(tasks, progress_callback)

    async def _collect_indexed_results(
        self,
        tasks: list[asyncio.Task[tuple[int, str, list[Any]]]],
        progress_callback: Callable[[int, int, str], None] | None,
    ) -> list[Any]:
        indexed_results: dict[int, list[Any]] = {}
        total = len(tasks)
        completed = 0
        for task in asyncio.as_completed(tasks):
            index, label, query_results = await task
            indexed_results[index] = query_results
            completed += 1
            if progress_callback is not None:
                progress_callback(completed, total, label)
        ordered_results = [indexed_results[index] for index in sorted(indexed_results)]
        return [item for group in ordered_results for item in group]

    async def _fetch_indexed_keyword_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        query: str,
        params: dict[str, Any],
    ) -> tuple[int, str, list[BukvarixKeywordResult]]:
        return index, query, await self._fetch_keyword_single(session, query, params)

    async def _fetch_indexed_domain_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        domain: str,
        params: dict[str, Any],
    ) -> tuple[int, str, list[BukvarixDomainResult]]:
        return index, domain, await self._fetch_domain_single(session, domain, params)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(min=1, max=10),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
        reraise=True,
    )
    async def _fetch_keyword_single(
        self,
        session: aiohttp.ClientSession,
        query: str,
        params: dict[str, Any],
    ) -> list[BukvarixKeywordResult]:
        request_params = self._build_params(query, params, COMMON_OPTION_NAMES + KEYWORD_OPTION_NAMES)
        return await self._request_keyword(session, "GET", BUKVARIX_KEYWORDS_ENDPOINT, query, request_params)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(min=1, max=10),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
        reraise=True,
    )
    async def _fetch_keyword_batch(
        self,
        session: aiohttp.ClientSession,
        queries: list[str],
        params: dict[str, Any],
    ) -> list[BukvarixKeywordResult]:
        query_label = f"{len(queries)} phrases"
        request_params = self._build_params("\r\n".join(queries), params, COMMON_OPTION_NAMES + KEYWORD_OPTION_NAMES)
        return await self._request_keyword(session, "POST", BUKVARIX_MKEYWORDS_ENDPOINT, query_label, request_params)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(min=1, max=10),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
        reraise=True,
    )
    async def _fetch_domain_single(
        self,
        session: aiohttp.ClientSession,
        domain: str,
        params: dict[str, Any],
    ) -> list[BukvarixDomainResult]:
        mode = str(params.get("mode", "single")).strip().lower()
        endpoint = BUKVARIX_SITE_CMP_ENDPOINT if mode == "compare" else BUKVARIX_SITE_ENDPOINT
        request_params = self._build_params(domain, params, COMMON_OPTION_NAMES + DOMAIN_OPTION_NAMES)
        return await self._request_domain(session, "GET", endpoint, domain, request_params)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(min=1, max=10),
        retry=retry_if_exception_type((aiohttp.ClientError, asyncio.TimeoutError)),
        reraise=True,
    )
    async def _fetch_domain_batch(
        self,
        session: aiohttp.ClientSession,
        domains: list[str],
        params: dict[str, Any],
    ) -> list[BukvarixDomainResult]:
        domain_label = f"{len(domains)} domains"
        request_params = self._build_params("\r\n".join(domains), params, COMMON_OPTION_NAMES + DOMAIN_OPTION_NAMES)
        return await self._request_domain(session, "POST", BUKVARIX_SITE_MCMP_ENDPOINT, domain_label, request_params)

    async def _request_keyword(
        self,
        session: aiohttp.ClientSession,
        method: str,
        endpoint: str,
        source_query: str,
        request_params: dict[str, Any],
    ) -> list[BukvarixKeywordResult]:
        async with self.semaphore:
            logger.info("Query sent | engine={} query={} page={}", "bukvarix_keywords", source_query, "")
            response_text, status = await self._request(session, method, endpoint, request_params)
        if status != 200:
            return [BukvarixKeywordResult(source_query=source_query, error_code=str(status), error_message=response_text[:300])]
        return self._parse_keyword_response(source_query, response_text, str(request_params.get("format", "txt")))

    async def _request_domain(
        self,
        session: aiohttp.ClientSession,
        method: str,
        endpoint: str,
        source_domain: str,
        request_params: dict[str, Any],
    ) -> list[BukvarixDomainResult]:
        async with self.semaphore:
            logger.info("Query sent | engine={} query={} page={}", "bukvarix_domains", source_domain, "")
            response_text, status = await self._request(session, method, endpoint, request_params)
        if status != 200:
            return [BukvarixDomainResult(source_domain=source_domain, error_code=str(status), error_message=response_text[:300])]
        return self._parse_domain_response(source_domain, response_text, str(request_params.get("format", "txt")))

    async def _request(
        self,
        session: aiohttp.ClientSession,
        method: str,
        endpoint: str,
        request_params: dict[str, Any],
    ) -> tuple[str, int]:
        if method == "POST":
            async with session.post(endpoint, data=request_params) as response:
                return await response.text(), response.status
        async with session.get(endpoint, params=request_params) as response:
            return await response.text(), response.status

    def _build_params(self, query: str, params: dict[str, Any], option_names: Iterable[str]) -> dict[str, Any]:
        request_params: dict[str, Any] = {
            "q": query,
            "api_key": self.api_key,
        }
        for option_name in option_names:
            value = params.get(option_name, "")
            if value != "":
                request_params[option_name] = value
        return request_params

    def _parse_keyword_response(
        self,
        source_query: str,
        response_text: str,
        response_format: str,
    ) -> list[BukvarixKeywordResult]:
        rows = self._parse_response_rows(response_text, response_format)
        if not rows:
            return [BukvarixKeywordResult(source_query=source_query, error_code="empty", error_message="Пустой ответ Bukvarix")]
        return [self._build_keyword_result(source_query, row) for row in rows]

    def _parse_domain_response(
        self,
        source_domain: str,
        response_text: str,
        response_format: str,
    ) -> list[BukvarixDomainResult]:
        rows = self._parse_response_rows(response_text, response_format)
        if not rows:
            return [BukvarixDomainResult(source_domain=source_domain, error_code="empty", error_message="Пустой ответ Bukvarix")]
        return [self._build_domain_result(source_domain, row) for row in rows]

    def _parse_response_rows(self, response_text: str, response_format: str) -> list[list[str]]:
        cleaned_text = response_text.lstrip("\ufeff").strip()
        if not cleaned_text:
            return []
        normalized_format = response_format.strip().lower()
        if normalized_format == "json":
            return self._parse_json_rows(cleaned_text)
        if normalized_format in {"csv", "tsv"}:
            delimiter = ";" if normalized_format == "csv" else "\t"
            rows = list(csv.reader(StringIO(cleaned_text), delimiter=delimiter))
        else:
            rows = [line.split("\t") if "\t" in line else [line] for line in cleaned_text.splitlines()]
        return [row for row in self._drop_header(rows) if any(cell.strip() for cell in row)]

    def _parse_json_rows(self, response_text: str) -> list[list[str]]:
        try:
            payload = json.loads(response_text)
        except json.JSONDecodeError:
            return [[response_text]]
        rows: list[list[str]] = []
        for item in self._iter_json_items(payload):
            if isinstance(item, dict):
                rows.append([str(value) for value in item.values()])
            elif isinstance(item, (list, tuple)):
                rows.append([str(value) for value in item])
            else:
                rows.append([str(item)])
        return rows

    def _iter_json_items(self, payload: Any) -> list[Any]:
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            for key in ("data", "rows", "items", "result", "results"):
                value = payload.get(key)
                if isinstance(value, list):
                    return value
            return [payload]
        return [payload]

    def _drop_header(self, rows: list[list[str]]) -> list[list[str]]:
        if not rows:
            return rows
        first_row_text = " ".join(rows[0]).casefold()
        header_markers = ("ключ", "слово", "фраз", "частот", "позици", "keyword", "words", "chars")
        if any(marker in first_row_text for marker in header_markers):
            return rows[1:]
        return rows

    def _build_keyword_result(self, source_query: str, row: list[str]) -> BukvarixKeywordResult:
        normalized_row = self._pad_row(row, 5)
        return BukvarixKeywordResult(
            source_query=source_query,
            keyword=normalized_row[0],
            words_count=normalized_row[1],
            chars_count=normalized_row[2],
            broad_frequency=normalized_row[3],
            exact_frequency=normalized_row[4],
            raw="\t".join(row),
        )

    def _build_domain_result(self, source_domain: str, row: list[str]) -> BukvarixDomainResult:
        normalized_row = self._pad_row(row, 7)
        return BukvarixDomainResult(
            source_domain=source_domain,
            keyword=normalized_row[0],
            words_count=normalized_row[1],
            chars_count=normalized_row[2],
            serp_results=normalized_row[3],
            broad_frequency=normalized_row[4],
            exact_frequency=normalized_row[5],
            position=normalized_row[6],
            raw="\t".join(row),
        )

    def _pad_row(self, row: list[str], size: int) -> list[str]:
        cleaned_row = [cell.strip() for cell in row]
        return cleaned_row + [""] * max(size - len(cleaned_row), 0)
