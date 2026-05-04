"""Клиент XMLRiver для получения XML-выдачи."""

from __future__ import annotations

import asyncio
import xml.etree.ElementTree as ET
from collections.abc import Callable
from typing import Any
from urllib.parse import urlsplit

import aiohttp
from aiohttp import ClientTimeout
from loguru import logger
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from app.models import XmlRiverDomainTopResult, XmlRiverResult

XMLRIVER_ENDPOINTS = {
    "yandex": "https://xmlriver.com/search_yandex/xml",
    "google": "https://xmlriver.com/search/xml",
}
DOMAIN_TOP_NOT_FOUND_POSITION = "Не найдено"


class XmlRiverClient:
    """Выполняет запросы к XMLRiver с ограничением параллельности."""

    def __init__(self, user_id: str, api_key: str, connect_timeout: int, read_timeout: int, max_concurrency: int) -> None:
        self.user_id = user_id
        self.api_key = api_key
        self.timeout = ClientTimeout(connect=connect_timeout, sock_read=read_timeout)
        self.semaphore = asyncio.Semaphore(max_concurrency)

    async def fetch_queries(
        self,
        queries: list[str],
        engine: str,
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[XmlRiverResult]:
        """Выполняет пакет запросов и возвращает плоский список результатов."""
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            tasks = [
                asyncio.create_task(self._fetch_indexed_query(session, index, query, engine, params))
                for index, query in enumerate(queries)
            ]
            indexed_results: dict[int, list[XmlRiverResult]] = {}
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

    async def fetch_domain_top_queries(
        self,
        queries: list[str],
        target_domain: str,
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[XmlRiverDomainTopResult]:
        """Checks whether the target domain appears in Yandex XMLRiver results for each query."""
        normalized_target_domain = normalize_domain(target_domain)
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            tasks = [
                asyncio.create_task(
                    self._fetch_indexed_domain_top_query(
                        session,
                        index,
                        query,
                        normalized_target_domain,
                        params,
                    ),
                )
                for index, query in enumerate(queries)
            ]
            indexed_results: dict[int, XmlRiverDomainTopResult] = {}
            total_queries = len(tasks)
            completed_queries = 0
            for task in asyncio.as_completed(tasks):
                index, query_result = await task
                indexed_results[index] = query_result
                completed_queries += 1
                if progress_callback is not None:
                    progress_callback(completed_queries, total_queries, query_result.query)
        return [indexed_results[index] for index in sorted(indexed_results)]

    async def _fetch_indexed_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        query: str,
        engine: str,
        params: dict[str, Any],
    ) -> tuple[int, list[XmlRiverResult]]:
        """Возвращает индекс исходного запроса вместе с результатами."""
        return index, await self._fetch_single_query(session, query, engine, params)

    async def _fetch_indexed_domain_top_query(
        self,
        session: aiohttp.ClientSession,
        index: int,
        query: str,
        normalized_target_domain: str,
        params: dict[str, Any],
    ) -> tuple[int, XmlRiverDomainTopResult]:
        """Returns the original query index and the domain-top check result."""
        query_results = await self._fetch_single_query(session, query, "yandex", params)
        return index, self._select_domain_top_result(query, normalized_target_domain, query_results)

    def _select_domain_top_result(
        self,
        query: str,
        normalized_target_domain: str,
        query_results: list[XmlRiverResult],
    ) -> XmlRiverDomainTopResult:
        """Finds the first Yandex result matching the target domain."""
        if not normalized_target_domain:
            return XmlRiverDomainTopResult(
                query=query,
                position="Ошибка",
                target_domain=normalized_target_domain,
                error_code="domain",
                error_message="Введите домен для проверки",
            )
        if not query_results:
            return XmlRiverDomainTopResult(
                query=query,
                position=DOMAIN_TOP_NOT_FOUND_POSITION,
                target_domain=normalized_target_domain,
            )
        first_result = query_results[0]
        if first_result.error_code == "empty":
            return XmlRiverDomainTopResult(
                query=query,
                position=DOMAIN_TOP_NOT_FOUND_POSITION,
                target_domain=normalized_target_domain,
            )
        if first_result.error_code:
            return XmlRiverDomainTopResult(
                query=query,
                position="Ошибка",
                url=first_result.url,
                domain=first_result.domain,
                target_domain=normalized_target_domain,
                error_code=first_result.error_code,
                error_message=first_result.error_message,
            )
        for result in query_results:
            result_domain = result.domain or result.url
            if domains_match(result_domain, normalized_target_domain):
                return XmlRiverDomainTopResult(
                    query=query,
                    position=result.position,
                    url=result.url,
                    domain=normalize_domain(result_domain),
                    target_domain=normalized_target_domain,
                )
        return XmlRiverDomainTopResult(
            query=query,
            position=DOMAIN_TOP_NOT_FOUND_POSITION,
            target_domain=normalized_target_domain,
        )

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
        engine: str,
        params: dict[str, Any],
    ) -> list[XmlRiverResult]:
        """Отправляет один запрос и парсит XML-ответ."""
        request_params = self._build_params(query, params)
        endpoint = XMLRIVER_ENDPOINTS[engine]
        async with self.semaphore:
            logger.info("Query sent | engine={} query={} page={}", engine, query, request_params.get("page", ""))
            async with session.get(endpoint, params=request_params) as response:
                xml_text = await response.text()
                if response.status != 200:
                    return [XmlRiverResult(query=query, error_code=str(response.status), error_message=xml_text[:300])]
        return self._parse_xml_results(query, xml_text)

    def _build_params(self, query: str, params: dict[str, Any]) -> dict[str, Any]:
        """Формирует GET-параметры XMLRiver."""
        engine = str(params.get("engine", "")).strip().lower()
        request_params = {
            "user": self.user_id,
            "key": self.api_key,
            "query": query,
            "groupby": params.get("groupby", "10"),
        }
        if "page" in params:
            request_params["page"] = params.get("page", "0")
        if engine == "google":
            request_params["loc"] = params.get("loc", "")
            request_params["country"] = params.get("country", "")
            request_params["lr"] = params.get("lr", "")
            request_params["domain"] = params.get("domain", "")
            request_params["device"] = params.get("device", "desktop")
        else:
            request_params["lr"] = params.get("lr", "1")
            request_params["domain"] = params.get("domain", "ru")
            request_params["lang"] = params.get("lang", "ru")
            request_params["device"] = params.get("device", "desktop")
            additional = str(params.get("additional", "")).strip()
            if additional:
                request_params["additional"] = additional
        request_params = {key: value for key, value in request_params.items() if value != ""}
        return request_params

    def _parse_xml_results(self, query: str, xml_text: str) -> list[XmlRiverResult]:
        """Преобразует XML в список строк результата."""
        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError as error:
            logger.error("API error | code={} query={} error={}", "xml_parse", query, str(error))
            return [XmlRiverResult(query=query, error_code="xml_parse", error_message=str(error))]

        api_error = self._read_api_error(root)
        if api_error is not None:
            error_code, error_message = api_error
            logger.error("API error | code={} query={} error={}", error_code, query, error_message)
            return [XmlRiverResult(query=query, error_code=error_code, error_message=error_message)]

        results: list[XmlRiverResult] = []
        for index, group in enumerate(root.findall(".//group"), start=1):
            doc = group.find("doc")
            if doc is None:
                continue
            position = self._read_text(doc, "position") or str(index)
            results.append(
                XmlRiverResult(
                    query=query,
                    position=position,
                    url=self._read_text(doc, "url"),
                    domain=self._read_text(doc, "domain"),
                    title=self._read_text(doc, "title"),
                    snippet=self._read_text(doc, "snippet") or self._read_first_passage(doc),
                ),
            )
        if results:
            return results
        return [XmlRiverResult(query=query, error_code="empty", error_message="Пустой ответ XMLRiver")]

    def _read_api_error(self, root: ET.Element) -> tuple[str, str] | None:
        """Reads a structured XMLRiver/Yandex XML error node if present."""
        error_node = root.find(".//error")
        if error_node is None:
            return None
        error_code = error_node.attrib.get("code", "").strip()
        error_message = (error_node.text or "").strip()
        code_node_text = self._read_text(error_node, "code")
        message_node_text = self._read_text(error_node, "message")
        return (
            error_code or code_node_text or "xmlriver_error",
            error_message or message_node_text or "XMLRiver returned an error",
        )

    def _read_first_passage(self, parent: ET.Element) -> str:
        """Returns first organic passage text when XMLRiver does not expose a snippet tag."""
        passage_node = parent.find("./passages/passage")
        if passage_node is None or passage_node.text is None:
            return ""
        return passage_node.text.strip()

    def _read_text(self, parent: ET.Element, tag_name: str) -> str:
        """Возвращает текст первого найденного тега или пустую строку."""
        node = parent.find(tag_name)
        if node is None or node.text is None:
            return ""
        return node.text.strip()


def normalize_domain(value: str) -> str:
    """Normalizes a domain or URL for comparison."""
    cleaned_value = value.strip().casefold()
    if not cleaned_value:
        return ""
    parse_value = cleaned_value if "://" in cleaned_value else f"//{cleaned_value}"
    parsed_value = urlsplit(parse_value)
    host = parsed_value.hostname or cleaned_value.split("/", 1)[0].split("?", 1)[0].split("#", 1)[0]
    host = host.split(":", 1)[0].strip().rstrip(".")
    if host.startswith("www."):
        host = host[4:]
    return host


def domains_match(candidate_domain: str, target_domain: str) -> bool:
    """Returns True when candidate equals target or is its subdomain."""
    normalized_candidate = normalize_domain(candidate_domain)
    normalized_target = normalize_domain(target_domain)
    if not normalized_candidate or not normalized_target:
        return False
    return normalized_candidate == normalized_target or normalized_candidate.endswith(f".{normalized_target}")
