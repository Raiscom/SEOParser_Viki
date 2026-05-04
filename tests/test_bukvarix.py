from __future__ import annotations

import unittest

from app.services.bukvarix import BukvarixClient


class BukvarixParsingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = BukvarixClient(api_key="free", connect_timeout=1, read_timeout=1, max_concurrency=1)

    def test_parses_keyword_tsv_with_header(self) -> None:
        response_text = (
            "Ключевое слово\tСлов\tСимволов\tШирокая частотность\tТочная частотность\n"
            "пластиковые окна\t2\t16\t1000\t100\n"
        )

        results = self.client._parse_keyword_response("окна", response_text, "tsv")

        self.assertEqual(results[0].source_query, "окна")
        self.assertEqual(results[0].keyword, "пластиковые окна")
        self.assertEqual(results[0].broad_frequency, "1000")
        self.assertEqual(results[0].exact_frequency, "100")

    def test_parses_domain_tsv_with_position(self) -> None:
        response_text = (
            "Ключевое слово\tСлов\tСимволов\tРезультатов\tШирокая\tТочная\tПозиция\n"
            "купить окна\t2\t10\t120000\t900\t90\t7\n"
        )

        results = self.client._parse_domain_response("example.ru", response_text, "tsv")

        self.assertEqual(results[0].source_domain, "example.ru")
        self.assertEqual(results[0].keyword, "купить окна")
        self.assertEqual(results[0].serp_results, "120000")
        self.assertEqual(results[0].position, "7")

    def test_build_params_omits_empty_options(self) -> None:
        params = self.client._build_params("окна", {"format": "tsv", "bom": "", "num": "250"}, ("format", "bom", "num"))

        self.assertEqual(params["q"], "окна")
        self.assertEqual(params["api_key"], "free")
        self.assertEqual(params["format"], "tsv")
        self.assertEqual(params["num"], "250")
        self.assertNotIn("bom", params)


if __name__ == "__main__":
    unittest.main()
