import asyncio
import json
import os
import sys
import unittest
from datetime import datetime


sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM, format_date


def _record(entry_id, subject, start_dt, end_dt, location="", organizer="", all_day=False):
    parts = [
        entry_id,
        subject,
        start_dt.strftime("%A, %d %B %Y at %H:%M:%S"),
        end_dt.strftime("%A, %d %B %Y at %H:%M:%S"),
        location,
        organizer,
        "true" if all_day else "false",
        str(start_dt.year),
        str(start_dt.month),
        str(start_dt.day),
        str(start_dt.hour),
        str(start_dt.minute),
        str(start_dt.second),
        str(end_dt.year),
        str(end_dt.month),
        str(end_dt.day),
        str(end_dt.hour),
        str(end_dt.minute),
        str(end_dt.second),
    ]
    return DELIM.join(parts)


class FakeBridge:
    def __init__(self, raw):
        self.raw = raw
        self.script = ""

    async def run(self, script):
        self.script = script
        return self.raw


class MacCalendarDateTests(unittest.TestCase):
    def setUp(self):
        self.original_bridge = server_mac.bridge

    def tearDown(self):
        server_mac.bridge = self.original_bridge

    def test_format_date_avoids_iso_like_applescript_dates(self):
        formatted = format_date(datetime(2026, 5, 3, 9, 30, 0))

        self.assertEqual(formatted, 'date "3 May 2026 at 09:30:00"')
        self.assertNotIn("2026-05-03", formatted)

    def test_list_events_filters_sorts_and_limits_synthetic_records(self):
        raw = RECORD_DELIM.join([
            _record(
                "later",
                "Later Synthetic Event",
                datetime(2026, 5, 10, 10, 0, 0),
                datetime(2026, 5, 10, 11, 0, 0),
            ),
            _record(
                "outside",
                "Outside Synthetic Event",
                datetime(2026, 5, 11, 9, 0, 0),
                datetime(2026, 5, 11, 10, 0, 0),
            ),
            _record(
                "earlier",
                "Earlier Synthetic Event",
                datetime(2026, 5, 4, 9, 0, 0),
                datetime(2026, 5, 4, 10, 0, 0),
            ),
        ])
        fake_bridge = FakeBridge(raw)
        server_mac.bridge = fake_bridge

        result = asyncio.run(server_mac.list_events("2026-05-03", "2026-05-10", 1))
        events = json.loads(result)

        self.assertEqual([event["entry_id"] for event in events], ["earlier"])
        self.assertIn('date "3 May 2026 at 00:00:00"', fake_bridge.script)
        self.assertIn('date "10 May 2026 at 23:59:59"', fake_bridge.script)

    def test_search_events_applies_query_and_range_synthetic_records(self):
        raw = RECORD_DELIM.join([
            _record(
                "match",
                "Planning Sync",
                datetime(2026, 5, 4, 9, 0, 0),
                datetime(2026, 5, 4, 10, 0, 0),
            ),
            _record(
                "query-miss",
                "Review",
                datetime(2026, 5, 5, 9, 0, 0),
                datetime(2026, 5, 5, 10, 0, 0),
            ),
            _record(
                "range-miss",
                "Planning Sync",
                datetime(2026, 5, 11, 9, 0, 0),
                datetime(2026, 5, 11, 10, 0, 0),
            ),
        ])
        fake_bridge = FakeBridge(raw)
        server_mac.bridge = fake_bridge

        result = asyncio.run(server_mac.search_events("planning", "2026-05-03", "2026-05-10", 5))
        events = json.loads(result)

        self.assertEqual([event["entry_id"] for event in events], ["match"])


if __name__ == "__main__":
    unittest.main()
