from __future__ import annotations

import asyncio
import unittest
from typing import Any

import aiohttp

from app.models import XmlRiverResult
from app.services.xmlriver import XmlRiverClient


class ControlledBatchXmlRiverClient(XmlRiverClient):
    def __init__(self, max_concurrency: int) -> None:
        super().__init__(
            user_id="user",
            api_key="key",
            connect_timeout=1,
            read_timeout=1,
            max_concurrency=max_concurrency,
        )
        self.started: list[int] = []
        self.batch_started_events = [asyncio.Event() for _ in range(3)]
        self.release_events = [asyncio.Event() for _ in range(3)]

    async def _fetch_single_query(
        self,
        session: aiohttp.ClientSession,
        query: str,
        engine: str,
        params: dict[str, Any],
    ) -> list[XmlRiverResult]:
        query_index = int(query.removeprefix("query-"))
        batch_index = query_index // self.batch_size
        self.started.append(query_index)
        self.batch_started_events[batch_index].set()
        await self.release_events[batch_index].wait()
        await asyncio.sleep((self.batch_size - query_index % self.batch_size) * 0.001)
        return [
            XmlRiverResult(
                query=query,
                position=str(query_index + 1),
                url=f"https://example.ru/{query_index}",
                domain="example.ru",
            ),
        ]


class XmlRiverBatchingTests(unittest.IsolatedAsyncioTestCase):
    async def test_fetch_queries_runs_completed_batches_and_preserves_order(self) -> None:
        client = ControlledBatchXmlRiverClient(max_concurrency=10)
        queries = [f"query-{index}" for index in range(25)]

        fetch_task = asyncio.create_task(client.fetch_queries(queries, "yandex", {"engine": "yandex"}))
        await client.batch_started_events[0].wait()
        await asyncio.sleep(0.01)

        self.assertEqual(client.started, list(range(10)))

        client.release_events[0].set()
        await client.batch_started_events[1].wait()
        await asyncio.sleep(0.01)

        self.assertEqual(client.started, list(range(20)))

        client.release_events[1].set()
        await client.batch_started_events[2].wait()
        await asyncio.sleep(0.01)

        self.assertEqual(client.started, list(range(25)))

        client.release_events[2].set()
        results = await fetch_task

        self.assertEqual([result.query for result in results], queries)
        self.assertEqual([result.position for result in results], [str(index + 1) for index in range(25)])

    async def test_fetch_queries_progress_counts_queries_not_batches(self) -> None:
        client = ControlledBatchXmlRiverClient(max_concurrency=10)
        queries = [f"query-{index}" for index in range(25)]
        progress: list[tuple[int, int]] = []

        fetch_task = asyncio.create_task(
            client.fetch_queries(
                queries,
                "yandex",
                {"engine": "yandex"},
                progress_callback=lambda completed, total, query: progress.append((completed, total)),
            ),
        )
        for batch_index in range(3):
            await client.batch_started_events[batch_index].wait()
            client.release_events[batch_index].set()
        await fetch_task

        self.assertEqual(progress[-1], (25, 25))
        self.assertEqual([completed for completed, total in progress], list(range(1, 26)))
        self.assertEqual({total for completed, total in progress}, {25})

    async def test_fetch_domain_top_queries_uses_same_batching(self) -> None:
        client = ControlledBatchXmlRiverClient(max_concurrency=10)
        queries = [f"query-{index}" for index in range(25)]

        fetch_task = asyncio.create_task(
            client.fetch_domain_top_queries(queries, "example.ru", {"engine": "yandex"}),
        )
        await client.batch_started_events[0].wait()
        await asyncio.sleep(0.01)

        self.assertEqual(client.started, list(range(10)))

        for batch_index in range(3):
            client.release_events[batch_index].set()
            if batch_index < 2:
                await client.batch_started_events[batch_index + 1].wait()
        results = await fetch_task

        self.assertEqual([result.query for result in results], queries)
        self.assertEqual([result.position for result in results], [str(index + 1) for index in range(25)])


if __name__ == "__main__":
    unittest.main()
