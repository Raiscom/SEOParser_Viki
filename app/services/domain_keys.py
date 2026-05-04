"""Domain keyword discovery through XMLRiver Yandex SERP."""

from __future__ import annotations

import asyncio
import html
import ipaddress
import json
import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Any
from urllib.parse import urlsplit

import aiohttp
from aiohttp import ClientTimeout
from loguru import logger

from app.models import DomainKeywordCandidate, DomainKeywordPosition, XmlRiverDomainTopResult
from app.services.xmlriver import XmlRiverClient, domains_match, normalize_domain

try:
    from bs4 import BeautifulSoup
except ImportError:  # pragma: no cover - exercised only without optional runtime dependency
    BeautifulSoup = None  # type: ignore[assignment]


YANDEX_TIPS_ENDPOINT = "https://xmlriver.com/search_yandex/xml"
DEFAULT_USER_AGENT = "SEOParserViki/1.0 (+https://xmlriver.com/)"
TOKEN_RE = re.compile(r"[0-9A-Za-zА-Яа-яЁё]+")
TAG_RE = re.compile(r"<[^>]+>")
WHITESPACE_RE = re.compile(r"\s+")
MAX_HTML_BYTES = 1_000_000
NOT_CHECKED_POSITION = "Не проверено"

STOP_WORDS = {
    "а",
    "без",
    "бы",
    "в",
    "вам",
    "вас",
    "ваш",
    "ваша",
    "ваше",
    "ваши",
    "во",
    "все",
    "для",
    "до",
    "его",
    "ее",
    "если",
    "есть",
    "и",
    "из",
    "или",
    "их",
    "к",
    "как",
    "ко",
    "на",
    "над",
    "не",
    "но",
    "о",
    "об",
    "от",
    "по",
    "под",
    "при",
    "с",
    "со",
    "так",
    "то",
    "у",
    "что",
    "это",
    "the",
    "and",
    "for",
    "from",
    "with",
    "into",
    "your",
    "you",
}


@dataclass(frozen=True)
class DomainKeywordFilters:
    """User-provided filters for URL and phrase collection."""

    include_words: tuple[str, ...] = ()
    exclude_words: tuple[str, ...] = ()
    url_include: tuple[str, ...] = ()
    url_exclude: tuple[str, ...] = ()


@dataclass(frozen=True)
class DomainKeywordOptions:
    """Collection limits and mode flags for domain keyword discovery."""

    max_urls: int = 50
    max_keywords: int = 200
    pages_per_seed: int = 1
    fetch_pages: bool = True
    use_tips: bool = True
    tips_seed_limit: int = 25


@dataclass
class PageSeoData:
    """SEO zones extracted from one URL or SERP row."""

    url: str = ""
    title: str = ""
    h1: str = ""
    h2: str = ""
    description: str = ""
    snippet: str = ""
    source: str = "serp"


@dataclass
class _KeywordScore:
    phrase: str
    score: float = 0.0
    source: str = ""
    source_url: str = ""
    title: str = ""
    h1: str = ""


class DomainKeysClient:
    """Collects probable domain keywords and checks their Yandex positions."""

    def __init__(
        self,
        user_id: str,
        api_key: str,
        connect_timeout: int,
        read_timeout: int,
        max_concurrency: int,
    ) -> None:
        self.user_id = user_id
        self.api_key = api_key
        self.timeout = ClientTimeout(connect=connect_timeout, sock_read=read_timeout)
        self.semaphore = asyncio.Semaphore(max_concurrency)
        self.xmlriver_client = XmlRiverClient(user_id, api_key, connect_timeout, read_timeout, max_concurrency)

    async def collect_keywords(
        self,
        domain: str,
        seed_phrases: list[str],
        params: dict[str, Any],
        filters: DomainKeywordFilters | None = None,
        options: DomainKeywordOptions | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[DomainKeywordCandidate]:
        """Collects keyword candidates from filtered site:domain SERP and optional page SEO zones."""
        normalized_domain = normalize_domain(domain)
        if not normalized_domain:
            raise ValueError("Введите домен для сбора ключей")
        seeds = dedupe_preserve_order(seed_phrases)
        if not seeds:
            raise ValueError("Введите хотя бы одну seed-фразу")

        filters = filters or DomainKeywordFilters()
        options = options or DomainKeywordOptions()
        options = clamp_options(options)

        serp_pages = await self._collect_serp_pages(
            normalized_domain,
            seeds,
            params,
            filters,
            options,
            progress_callback,
        )
        page_data = list(serp_pages)
        if options.fetch_pages and serp_pages:
            fetched_pages = await self._fetch_pages(
                serp_pages[: options.max_urls],
                normalized_domain,
                filters,
                progress_callback,
            )
            page_data.extend(fetched_pages)

        candidates = generate_keyword_candidates(page_data, filters, options.max_keywords)
        if options.use_tips and candidates:
            tip_phrases = await self._fetch_tips(
                [candidate.phrase for candidate in candidates[: options.tips_seed_limit]],
                params,
                progress_callback,
            )
            candidates = merge_tip_candidates(candidates, tip_phrases, filters, options.max_keywords)
        return candidates[: options.max_keywords]

    async def check_positions(
        self,
        domain: str,
        candidates: list[DomainKeywordCandidate],
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[DomainKeywordPosition]:
        """Checks generated keywords through the existing XMLRiver domain-top logic."""
        queries = [candidate.phrase for candidate in candidates]
        candidate_by_phrase = {candidate.phrase: candidate for candidate in candidates}
        checked = await self.xmlriver_client.fetch_domain_top_queries(
            queries=queries,
            target_domain=domain,
            params=params,
            progress_callback=progress_callback,
        )
        return [merge_candidate_with_position(result, candidate_by_phrase.get(result.query)) for result in checked]

    async def collect_and_check(
        self,
        domain: str,
        seed_phrases: list[str],
        params: dict[str, Any],
        filters: DomainKeywordFilters | None = None,
        options: DomainKeywordOptions | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[DomainKeywordPosition]:
        """Collects keyword candidates and immediately checks their Yandex positions."""
        candidates = await self.collect_keywords(domain, seed_phrases, params, filters, options, progress_callback)
        return await self.check_positions(domain, candidates, params, progress_callback)

    async def _collect_serp_pages(
        self,
        domain: str,
        seed_phrases: list[str],
        params: dict[str, Any],
        filters: DomainKeywordFilters,
        options: DomainKeywordOptions,
        progress_callback: Callable[[int, int, str], None] | None,
    ) -> list[PageSeoData]:
        query_params = dict(params)
        query_params["engine"] = "yandex"
        query_params["groupby"] = "10"
        total = len(seed_phrases) * max(options.pages_per_seed, 1)
        completed = 0
        pages_by_url: dict[str, PageSeoData] = {}
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            for seed in seed_phrases:
                for page in range(options.pages_per_seed):
                    request_params = dict(query_params)
                    request_params["page"] = str(page)
                    query = build_site_query(domain, seed)
                    results = await self.xmlriver_client._fetch_single_query(session, query, "yandex", request_params)
                    for result in results:
                        if result.error_code:
                            logger.warning(
                                "Domain keys SERP row skipped | query={} code={} message={}",
                                query,
                                result.error_code,
                                result.error_message,
                            )
                            continue
                        if not is_allowed_domain_url(result.url, domain):
                            continue
                        if not url_passes_filters(result.url, filters):
                            continue
                        normalized_url = normalize_url_key(result.url)
                        if normalized_url not in pages_by_url:
                            pages_by_url[normalized_url] = PageSeoData(
                                url=result.url,
                                title=result.title,
                                snippet=result.snippet,
                                source="serp",
                            )
                        if len(pages_by_url) >= options.max_urls:
                            break
                    completed += 1
                    if progress_callback is not None:
                        progress_callback(completed, total, query)
                    if len(pages_by_url) >= options.max_urls:
                        break
                if len(pages_by_url) >= options.max_urls:
                    break
        return list(pages_by_url.values())

    async def _fetch_pages(
        self,
        pages: list[PageSeoData],
        domain: str,
        filters: DomainKeywordFilters,
        progress_callback: Callable[[int, int, str], None] | None,
    ) -> list[PageSeoData]:
        async with aiohttp.ClientSession(timeout=self.timeout, headers={"User-Agent": DEFAULT_USER_AGENT}) as session:
            tasks = [asyncio.create_task(self._fetch_page(session, page, domain, filters)) for page in pages]
            fetched: list[PageSeoData] = []
            total = len(tasks)
            completed = 0
            for task in asyncio.as_completed(tasks):
                page_data = await task
                if page_data is not None:
                    fetched.append(page_data)
                completed += 1
                if progress_callback is not None:
                    progress_callback(completed, total, page_data.url if page_data else "")
        return fetched

    async def _fetch_page(
        self,
        session: aiohttp.ClientSession,
        page: PageSeoData,
        domain: str,
        filters: DomainKeywordFilters,
    ) -> PageSeoData | None:
        if not is_allowed_domain_url(page.url, domain) or not url_passes_filters(page.url, filters):
            return None
        try:
            async with self.semaphore:
                async with session.get(page.url, allow_redirects=True, max_redirects=5) as response:
                    final_url = str(response.url)
                    if (
                        response.status >= 400
                        or not is_allowed_domain_url(final_url, domain)
                        or not url_passes_filters(final_url, filters)
                    ):
                        return None
                    content_type = response.headers.get("Content-Type", "").lower()
                    if content_type and "html" not in content_type:
                        return None
                    raw = await response.content.read(MAX_HTML_BYTES + 1)
        except (aiohttp.ClientError, asyncio.TimeoutError, UnicodeError) as error:
            logger.warning("Domain page fetch skipped | url={} error={}", page.url, str(error))
            return None
        if len(raw) > MAX_HTML_BYTES:
            raw = raw[:MAX_HTML_BYTES]
        html_text = raw.decode(detect_charset(raw), errors="ignore")
        extracted = extract_page_seo_data(html_text)
        extracted.url = str(response.url)
        extracted.snippet = page.snippet
        extracted.source = "page"
        return extracted

    async def _fetch_tips(
        self,
        phrases: list[str],
        params: dict[str, Any],
        progress_callback: Callable[[int, int, str], None] | None,
    ) -> list[str]:
        phrases = dedupe_preserve_order(phrases)[:50]
        if not phrases:
            return []
        request_params = {
            "setab": "tips",
            "user": self.user_id,
            "key": self.api_key,
            "domain": params.get("domain", "ru"),
            "lr": params.get("lr", "1"),
        }
        async with aiohttp.ClientSession(timeout=self.timeout) as session:
            try:
                async with self.semaphore:
                    async with session.post(YANDEX_TIPS_ENDPOINT, params=request_params, json={"phrases": phrases}) as response:
                        response_text = await response.text()
            except (aiohttp.ClientError, asyncio.TimeoutError) as error:
                logger.warning("XMLRiver tips skipped | error={}", str(error))
                return []
        if progress_callback is not None:
            progress_callback(1, 1, "tips")
        if response.status != 200:
            logger.warning("XMLRiver tips skipped | status={} preview={}", response.status, response_text[:200])
            return []
        try:
            payload = json.loads(response_text)
        except json.JSONDecodeError:
            logger.warning("XMLRiver tips skipped | invalid_json preview={}", response_text[:200])
            return []
        tips = payload.get("phrases") if isinstance(payload, dict) else None
        if not isinstance(tips, list):
            return []
        return [str(item).strip() for item in tips if str(item).strip()]


def build_site_query(domain: str, seed_phrase: str) -> str:
    """Builds the filtered Yandex site: query used for URL discovery."""
    cleaned_domain = normalize_domain(domain)
    cleaned_seed = WHITESPACE_RE.sub(" ", seed_phrase.strip())
    return f"site:{cleaned_domain} {cleaned_seed}".strip()


def parse_filter_text(value: str) -> tuple[str, ...]:
    """Splits comma/newline separated filter text into normalized chunks."""
    chunks = re.split(r"[\n,;]+", value)
    return tuple(chunk.strip().casefold() for chunk in chunks if chunk.strip())


def clamp_options(options: DomainKeywordOptions) -> DomainKeywordOptions:
    """Bounds user limits to keep collection predictable."""
    return DomainKeywordOptions(
        max_urls=max(1, min(options.max_urls, 500)),
        max_keywords=max(1, min(options.max_keywords, 5000)),
        pages_per_seed=max(1, min(options.pages_per_seed, 5)),
        fetch_pages=options.fetch_pages,
        use_tips=options.use_tips,
        tips_seed_limit=max(1, min(options.tips_seed_limit, 50)),
    )


def is_allowed_domain_url(url: str, target_domain: str) -> bool:
    """Allows only safe http(s) URLs from the target domain or its subdomains."""
    parsed_url = urlsplit(url.strip())
    if parsed_url.scheme not in {"http", "https"}:
        return False
    host = parsed_url.hostname or ""
    if not host or is_private_or_local_host(host):
        return False
    return domains_match(host, target_domain)


def is_private_or_local_host(host: str) -> bool:
    """Rejects localhost and private IP hosts before fetching a URL."""
    cleaned_host = host.strip().strip("[]").casefold()
    if cleaned_host in {"localhost", "localhost.localdomain"} or cleaned_host.endswith(".local"):
        return True
    try:
        address = ipaddress.ip_address(cleaned_host)
    except ValueError:
        return False
    return any(
        [
            address.is_private,
            address.is_loopback,
            address.is_link_local,
            address.is_multicast,
            address.is_reserved,
            address.is_unspecified,
        ],
    )


def url_passes_filters(url: str, filters: DomainKeywordFilters) -> bool:
    """Applies URL include/exclude substring filters."""
    lowered_url = url.casefold()
    if filters.url_include and not any(item in lowered_url for item in filters.url_include):
        return False
    return not any(item in lowered_url for item in filters.url_exclude)


def phrase_passes_filters(phrase: str, filters: DomainKeywordFilters) -> bool:
    """Applies phrase include/exclude filters."""
    lowered_phrase = phrase.casefold()
    if filters.include_words and not any(item in lowered_phrase for item in filters.include_words):
        return False
    return not any(item in lowered_phrase for item in filters.exclude_words)


def extract_page_seo_data(html_text: str) -> PageSeoData:
    """Extracts title, H1, H2 and meta description without reading full body text."""
    if BeautifulSoup is not None:
        soup = BeautifulSoup(html_text, "html.parser")
        title = clean_text(soup.title.get_text(" ", strip=True) if soup.title else "")
        h1 = clean_text(" ".join(node.get_text(" ", strip=True) for node in soup.find_all("h1", limit=3)))
        h2 = clean_text(" ".join(node.get_text(" ", strip=True) for node in soup.find_all("h2", limit=8)))
        meta = soup.find("meta", attrs={"name": re.compile("^description$", re.I)})
        description = clean_text(str(meta.get("content", "")) if meta else "")
        return PageSeoData(title=title, h1=h1, h2=h2, description=description, source="page")
    parser = _SeoZoneParser()
    parser.feed(html_text)
    return PageSeoData(
        title=clean_text(parser.title),
        h1=clean_text(" ".join(parser.h1)),
        h2=clean_text(" ".join(parser.h2)),
        description=clean_text(parser.description),
        source="page",
    )


class _SeoZoneParser(HTMLParser):
    """Small fallback SEO-zone parser for environments without BeautifulSoup."""

    def __init__(self) -> None:
        super().__init__()
        self.current_tag = ""
        self.title = ""
        self.h1: list[str] = []
        self.h2: list[str] = []
        self.description = ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.current_tag = tag.lower()
        if self.current_tag == "meta":
            attr_map = {name.lower(): value or "" for name, value in attrs}
            if attr_map.get("name", "").casefold() == "description":
                self.description = attr_map.get("content", "")

    def handle_endtag(self, tag: str) -> None:
        if self.current_tag == tag.lower():
            self.current_tag = ""

    def handle_data(self, data: str) -> None:
        if self.current_tag == "title":
            self.title += f" {data}"
        elif self.current_tag == "h1" and len(self.h1) < 3:
            self.h1.append(data)
        elif self.current_tag == "h2" and len(self.h2) < 8:
            self.h2.append(data)


def generate_keyword_candidates(
    pages: Iterable[PageSeoData],
    filters: DomainKeywordFilters,
    max_keywords: int,
) -> list[DomainKeywordCandidate]:
    """Generates scored 2-5 word keyword candidates from SEO zones."""
    scores: dict[str, _KeywordScore] = {}
    weights = {
        "title": 5.0,
        "h1": 5.0,
        "h2": 3.0,
        "description": 3.0,
        "snippet": 2.0,
    }
    for page in pages:
        for zone_name, weight in weights.items():
            text = getattr(page, zone_name)
            if not text:
                continue
            for phrase in iter_ngrams(text):
                if not phrase_passes_filters(phrase, filters):
                    continue
                item = scores.get(phrase)
                if item is None:
                    scores[phrase] = _KeywordScore(
                        phrase=phrase,
                        score=weight,
                        source=page.source,
                        source_url=page.url,
                        title=page.title,
                        h1=page.h1,
                    )
                else:
                    item.score += weight
                    if not item.source_url and page.url:
                        item.source_url = page.url
                    if not item.title and page.title:
                        item.title = page.title
                    if not item.h1 and page.h1:
                        item.h1 = page.h1
                    if item.source != page.source:
                        item.source = "mixed"
    candidates = sorted(scores.values(), key=lambda item: (-item.score, item.phrase))
    return [
        DomainKeywordCandidate(
            phrase=item.phrase,
            source=item.source,
            score=round(item.score, 2),
            source_url=item.source_url,
            title=item.title,
            h1=item.h1,
        )
        for item in candidates[:max_keywords]
    ]


def merge_tip_candidates(
    candidates: list[DomainKeywordCandidate],
    tip_phrases: list[str],
    filters: DomainKeywordFilters,
    max_keywords: int,
) -> list[DomainKeywordCandidate]:
    """Adds XMLRiver tips to existing generated candidates."""
    by_phrase = {candidate.phrase.casefold(): candidate for candidate in candidates}
    for tip in tip_phrases:
        phrase = clean_phrase(tip)
        if not phrase or not phrase_passes_filters(phrase, filters):
            continue
        lowered = phrase.casefold()
        if lowered in by_phrase:
            current = by_phrase[lowered]
            current.score = round(current.score + 1.0, 2)
            if current.source != "tips" and "tips" not in current.source.split("+"):
                current.source = f"{current.source}+tips" if current.source else "tips"
            continue
        by_phrase[lowered] = DomainKeywordCandidate(phrase=phrase, source="tips", score=1.0)
    return sorted(by_phrase.values(), key=lambda item: (-item.score, item.phrase))[:max_keywords]


def merge_candidate_with_position(
    result: XmlRiverDomainTopResult,
    candidate: DomainKeywordCandidate | None,
) -> DomainKeywordPosition:
    """Combines a generated candidate with an XMLRiver domain-top result."""
    candidate = candidate or DomainKeywordCandidate(phrase=result.query)
    return DomainKeywordPosition(
        phrase=result.query,
        position=result.position,
        url=result.url,
        domain=result.domain,
        source=candidate.source,
        score=candidate.score,
        source_url=candidate.source_url,
        title=candidate.title,
        h1=candidate.h1,
        error_message=result.error_message,
        error_code=result.error_code,
    )


def candidates_to_unchecked_positions(candidates: list[DomainKeywordCandidate]) -> list[DomainKeywordPosition]:
    """Converts collected candidates to table rows before position checks."""
    return [
        DomainKeywordPosition(
            phrase=candidate.phrase,
            position=NOT_CHECKED_POSITION,
            source=candidate.source,
            score=candidate.score,
            source_url=candidate.source_url,
            title=candidate.title,
            h1=candidate.h1,
        )
        for candidate in candidates
    ]


def iter_ngrams(text: str, min_size: int = 2, max_size: int = 5) -> Iterable[str]:
    """Yields normalized n-grams from text."""
    tokens = tokenize(text)
    for size in range(min_size, max_size + 1):
        if len(tokens) < size:
            continue
        for index in range(0, len(tokens) - size + 1):
            chunk = tokens[index : index + size]
            if not is_good_ngram(chunk):
                continue
            yield " ".join(chunk)


def tokenize(text: str) -> list[str]:
    """Extracts normalized text tokens."""
    cleaned = clean_text(text).casefold()
    return [token for token in TOKEN_RE.findall(cleaned) if len(token) > 1]


def is_good_ngram(tokens: list[str]) -> bool:
    """Rejects n-grams that are mostly stop words or numeric noise."""
    if not tokens:
        return False
    if tokens[0] in STOP_WORDS or tokens[-1] in STOP_WORDS:
        return False
    meaningful = [token for token in tokens if token not in STOP_WORDS and not token.isdigit()]
    return len(meaningful) >= max(1, len(tokens) - 1)


def clean_phrase(text: str) -> str:
    """Normalizes a phrase for display and de-duplication."""
    return " ".join(tokenize(text))


def clean_text(text: str) -> str:
    """Removes tags/entities and collapses whitespace."""
    unescaped = html.unescape(TAG_RE.sub(" ", text))
    return WHITESPACE_RE.sub(" ", unescaped).strip()


def normalize_url_key(url: str) -> str:
    """Normalizes URLs for de-duplication while preserving path/query."""
    parsed = urlsplit(url.strip())
    host = (parsed.hostname or "").casefold()
    path = parsed.path.rstrip("/") or "/"
    query = f"?{parsed.query}" if parsed.query else ""
    return f"{parsed.scheme.casefold()}://{host}{path}{query}"


def dedupe_preserve_order(values: Iterable[str]) -> list[str]:
    """Deduplicates non-empty strings while preserving input order."""
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        cleaned = WHITESPACE_RE.sub(" ", value.strip())
        lowered = cleaned.casefold()
        if not cleaned or lowered in seen:
            continue
        seen.add(lowered)
        result.append(cleaned)
    return result


def detect_charset(raw: bytes) -> str:
    """Reads a basic HTML charset declaration."""
    head = raw[:2048].decode("ascii", errors="ignore")
    match = re.search(r"charset=[\"']?([A-Za-z0-9_\-]+)", head, flags=re.I)
    return match.group(1) if match else "utf-8"
