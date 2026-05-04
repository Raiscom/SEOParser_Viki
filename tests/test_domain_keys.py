from __future__ import annotations

import unittest

from app.services.domain_keys import (
    DomainKeywordFilters,
    PageSeoData,
    build_site_query,
    extract_page_seo_data,
    generate_keyword_candidates,
    is_allowed_domain_url,
    iter_ngrams,
    parse_filter_text,
    phrase_passes_filters,
    url_passes_filters,
)
from app.services.xmlriver import XmlRiverClient


class DomainKeysQueryTests(unittest.TestCase):
    def test_builds_site_query_from_domain_and_seed(self) -> None:
        self.assertEqual(build_site_query("https://www.Example.ru/path", " ремонт окон "), "site:example.ru ремонт окон")

    def test_parse_filter_text_splits_common_separators(self) -> None:
        self.assertEqual(parse_filter_text("окна, двери\nмонтаж"), ("окна", "двери", "монтаж"))


class DomainKeysSecurityTests(unittest.TestCase):
    def test_allows_target_domain_and_subdomains_only(self) -> None:
        self.assertTrue(is_allowed_domain_url("https://example.ru/catalog", "example.ru"))
        self.assertTrue(is_allowed_domain_url("https://blog.example.ru/post", "example.ru"))
        self.assertFalse(is_allowed_domain_url("https://badexample.ru/page", "example.ru"))
        self.assertFalse(is_allowed_domain_url("ftp://example.ru/file", "example.ru"))

    def test_blocks_private_and_local_hosts(self) -> None:
        self.assertFalse(is_allowed_domain_url("http://127.0.0.1/page", "127.0.0.1"))
        self.assertFalse(is_allowed_domain_url("http://localhost/page", "localhost"))
        self.assertFalse(is_allowed_domain_url("http://10.0.0.1/page", "10.0.0.1"))

    def test_applies_url_and_phrase_filters(self) -> None:
        filters = DomainKeywordFilters(
            include_words=("окна",),
            exclude_words=("бесплатно",),
            url_include=("/catalog",),
            url_exclude=("private",),
        )
        self.assertTrue(url_passes_filters("https://example.ru/catalog/okna", filters))
        self.assertFalse(url_passes_filters("https://example.ru/private/okna", filters))
        self.assertTrue(phrase_passes_filters("пластиковые окна", filters))
        self.assertFalse(phrase_passes_filters("бесплатно пластиковые окна", filters))
        self.assertFalse(phrase_passes_filters("пластиковые двери", filters))

    def test_positive_phrase_filters_match_any_item(self) -> None:
        filters = DomainKeywordFilters(include_words=("окна", "двери"))

        self.assertTrue(phrase_passes_filters("межкомнатные двери", filters))
        self.assertFalse(phrase_passes_filters("натяжные потолки", filters))


class DomainKeysExtractionTests(unittest.TestCase):
    def test_extracts_limited_seo_zones(self) -> None:
        html = """
        <html>
          <head>
            <title>Пластиковые окна в Москве</title>
            <meta name="description" content="Окна с установкой и гарантией">
          </head>
          <body>
            <nav>Меню футер мусор</nav>
            <h1>Купить пластиковые окна</h1>
            <h2>Монтаж окон под ключ</h2>
            <p>Этот длинный body-текст не должен попадать напрямую.</p>
          </body>
        </html>
        """
        data = extract_page_seo_data(html)

        self.assertEqual(data.title, "Пластиковые окна в Москве")
        self.assertEqual(data.h1, "Купить пластиковые окна")
        self.assertIn("Монтаж окон", data.h2)
        self.assertNotIn("длинный body", " ".join([data.title, data.h1, data.h2, data.description]))

    def test_generates_scored_candidates_from_seo_zones(self) -> None:
        pages = [
            PageSeoData(
                url="https://example.ru/okna",
                title="Пластиковые окна в Москве",
                h1="Купить пластиковые окна",
                snippet="Окна с установкой",
                source="page",
            ),
        ]
        candidates = generate_keyword_candidates(pages, DomainKeywordFilters(), 10)
        phrases = {candidate.phrase for candidate in candidates}

        self.assertIn("пластиковые окна", phrases)
        self.assertIn("купить пластиковые окна", phrases)

    def test_iter_ngrams_filters_stopword_edges(self) -> None:
        phrases = set(iter_ngrams("для пластиковые окна в Москве"))

        self.assertIn("пластиковые окна", phrases)
        self.assertNotIn("для пластиковые", phrases)


class XmlRiverErrorParsingTests(unittest.TestCase):
    def test_parses_xmlriver_error_node(self) -> None:
        client = XmlRiverClient("user", "key", 1, 1, 1)
        results = client._parse_xml_results("query", '<yandexsearch><error code="15">No money</error></yandexsearch>')

        self.assertEqual(results[0].error_code, "15")
        self.assertEqual(results[0].error_message, "No money")

    def test_reads_passage_as_snippet_fallback(self) -> None:
        client = XmlRiverClient("user", "key", 1, 1, 1)
        xml_text = """
        <yandexsearch>
          <response>
            <results>
              <grouping>
                <group>
                  <doc>
                    <url>https://example.ru/</url>
                    <domain>example.ru</domain>
                    <title>Example</title>
                    <passages><passage>Snippet from passage</passage></passages>
                  </doc>
                </group>
              </grouping>
            </results>
          </response>
        </yandexsearch>
        """
        results = client._parse_xml_results("query", xml_text)

        self.assertEqual(results[0].snippet, "Snippet from passage")


if __name__ == "__main__":
    unittest.main()
