from __future__ import annotations

import unittest

from app.services.xmlriver import DOMAIN_TOP_NOT_FOUND_POSITION, XmlRiverClient, domains_match, normalize_domain


class DomainMatchingTests(unittest.TestCase):
    def test_normalize_domain_from_url(self) -> None:
        self.assertEqual(normalize_domain("https://www.Example.ru:443/path?q=1"), "example.ru")

    def test_domain_matches_exact_www_and_subdomain(self) -> None:
        self.assertTrue(domains_match("example.ru", "example.ru"))
        self.assertTrue(domains_match("www.example.ru", "example.ru"))
        self.assertTrue(domains_match("blog.example.ru", "example.ru"))

    def test_domain_does_not_match_suffix_without_boundary(self) -> None:
        self.assertFalse(domains_match("badexample.ru", "example.ru"))
        self.assertFalse(domains_match("example.com", "example.ru"))


class DomainTopSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = XmlRiverClient(
            user_id="user",
            api_key="key",
            connect_timeout=1,
            read_timeout=1,
            max_concurrency=1,
        )

    def test_selects_first_matching_domain_from_xml_results(self) -> None:
        xml_text = """
        <yandexsearch>
          <response>
            <results>
              <grouping>
                <group>
                  <doc>
                    <position>1</position>
                    <url>https://other.ru/page</url>
                    <domain>other.ru</domain>
                    <title>Other</title>
                    <snippet>Other snippet</snippet>
                  </doc>
                </group>
                <group>
                  <doc>
                    <position>2</position>
                    <url>https://blog.example.ru/post</url>
                    <domain>blog.example.ru</domain>
                    <title>Example</title>
                    <snippet>Example snippet</snippet>
                  </doc>
                </group>
              </grouping>
            </results>
          </response>
        </yandexsearch>
        """
        parsed_results = self.client._parse_xml_results("query", xml_text)
        result = self.client._select_domain_top_result("query", "example.ru", parsed_results)

        self.assertEqual(result.position, "2")
        self.assertEqual(result.url, "https://blog.example.ru/post")
        self.assertEqual(result.domain, "blog.example.ru")
        self.assertEqual(result.error_code, "")

    def test_returns_not_found_when_domain_is_absent(self) -> None:
        xml_text = """
        <yandexsearch>
          <response>
            <results>
              <grouping>
                <group>
                  <doc>
                    <position>1</position>
                    <url>https://other.ru/page</url>
                    <domain>other.ru</domain>
                  </doc>
                </group>
              </grouping>
            </results>
          </response>
        </yandexsearch>
        """
        parsed_results = self.client._parse_xml_results("query", xml_text)
        result = self.client._select_domain_top_result("query", "example.ru", parsed_results)

        self.assertEqual(result.position, DOMAIN_TOP_NOT_FOUND_POSITION)
        self.assertEqual(result.error_code, "")

    def test_empty_xmlriver_results_are_not_found(self) -> None:
        parsed_results = self.client._parse_xml_results("query", "<yandexsearch></yandexsearch>")
        result = self.client._select_domain_top_result("query", "example.ru", parsed_results)

        self.assertEqual(result.position, DOMAIN_TOP_NOT_FOUND_POSITION)
        self.assertEqual(result.error_code, "")

    def test_preserves_xml_error_as_error_row(self) -> None:
        parsed_results = self.client._parse_xml_results("query", "<broken")
        result = self.client._select_domain_top_result("query", "example.ru", parsed_results)

        self.assertEqual(result.position, "Ошибка")
        self.assertEqual(result.error_code, "xml_parse")


if __name__ == "__main__":
    unittest.main()
