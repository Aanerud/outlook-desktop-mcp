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
    """Fake AppleScript bridge.

    ``_gather_events`` issues two queries per call: the non-recurring
    calendar query, then the recurring-master query. This fake returns
    ``raw`` for the calendar query and ``recurring_raw`` (default empty) for
    the recurring query, and records every script it was handed.
    """

    def __init__(self, raw, recurring_raw=""):
        self.raw = raw
        self.recurring_raw = recurring_raw
        self.scripts = []
        self.script = ""

    async def run(self, script):
        self.script = script
        self.scripts.append(script)
        if "is recurring is true" in script:
            return self.recurring_raw
        return self.raw

    @property
    def calendar_script(self):
        for s in self.scripts:
            if "is recurring is false" in s:
                return s
        return self.script


class MacCalendarDateTests(unittest.TestCase):
    def setUp(self):
        self.original_bridge = server_mac.bridge

    def tearDown(self):
        server_mac.bridge = self.original_bridge

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
        # Assert the range bounds are interpolated via format_date, independent of
        # its exact AppleScript encoding (see format_date_test.py for that).
        self.assertIn(format_date(datetime(2026, 5, 3, 0, 0, 0)), fake_bridge.calendar_script)
        self.assertIn(format_date(datetime(2026, 5, 10, 23, 59, 59)), fake_bridge.calendar_script)
        # The calendar query must exclude recurring masters (they are expanded
        # separately) to avoid double-counting.
        self.assertIn("is recurring is false", fake_bridge.calendar_script)

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


def _recurring_record(entry_id, subject, start_dt, end_dt, rtype, interval,
                      weekdays, until=None, count=None, all_day=False):
    """Build a recurring-master record in the layout emitted by
    _recurring_events_script(): 19 base fields + rtype, interval, weekday
    flags, an end marker, and end year/month/day + occurrence count."""
    wd = ",".join("true" if d else "false" for d in weekdays)
    uy, um, ud = ("", "", "")
    if until is not None:
        uy, um, ud = str(until.year), str(until.month), str(until.day)
    parts = [
        entry_id, subject,
        start_dt.strftime("%A, %d %B %Y at %H:%M:%S"),
        end_dt.strftime("%A, %d %B %Y at %H:%M:%S"),
        "", "", "true" if all_day else "false",
        str(start_dt.year), str(start_dt.month), str(start_dt.day),
        str(start_dt.hour), str(start_dt.minute), str(start_dt.second),
        str(end_dt.year), str(end_dt.month), str(end_dt.day),
        str(end_dt.hour), str(end_dt.minute), str(end_dt.second),
        rtype, str(interval), wd, "end",
        uy, um, ud, "" if count is None else str(count),
    ]
    return DELIM.join(parts)


class MacRecurringExpansionTests(unittest.TestCase):
    def setUp(self):
        self.original_bridge = server_mac.bridge

    def tearDown(self):
        server_mac.bridge = self.original_bridge

    def test_biweekly_master_is_expanded_into_range(self):
        # A biweekly Thursday series whose master start (first occurrence) is
        # well before the query range — the exact bug: it must still surface.
        recurring = _recurring_record(
            "retro", "Launch Retro",
            datetime(2025, 1, 30, 11, 0, 0), datetime(2025, 1, 30, 11, 45, 0),
            "weekly", 2, [False, False, False, True, False, False, False],
            until=datetime(2027, 8, 24),
        )
        server_mac.bridge = FakeBridge(raw="", recurring_raw=recurring)

        result = asyncio.run(server_mac.list_events("2026-08-27", "2026-08-27", 20))
        events = json.loads(result)

        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["subject"], "Launch Retro")
        self.assertTrue(events[0]["start"].startswith("Thursday, August 27, 2026"))
        self.assertIn("11:00:00 AM", events[0]["start"])
        self.assertEqual(events[0]["duration"], 45)

    def test_ended_series_and_off_week_are_excluded(self):
        # Series that ended in 2024 must not appear in 2026; and a biweekly
        # series must not appear on its off week.
        ended = _recurring_record(
            "old", "Old Retro",
            datetime(2022, 1, 6, 9, 0, 0), datetime(2022, 1, 6, 10, 0, 0),
            "weekly", 2, [False, False, False, True, False, False, False],
            until=datetime(2024, 8, 29),
        )
        biweekly = _recurring_record(
            "retro", "Launch Retro",
            datetime(2025, 1, 30, 11, 0, 0), datetime(2025, 1, 30, 11, 45, 0),
            "weekly", 2, [False, False, False, True, False, False, False],
            until=datetime(2027, 8, 24),
        )
        server_mac.bridge = FakeBridge(
            raw="", recurring_raw=RECORD_DELIM.join([ended, biweekly]))

        # 2026-09-03 is an off week for the biweekly series (and after the
        # ended series). Expect nothing.
        result = asyncio.run(server_mac.list_events("2026-09-03", "2026-09-03", 20))
        self.assertEqual(json.loads(result), [])

    def test_relative_monthly_weekday_reminder_does_not_leak_as_weekly(self):
        # A weekday-flagged "relative monthly" reminder must not be treated as
        # a weekly series, or it would appear on many wrong days.
        reminder = _recurring_record(
            "pto", "Schedule PTO for this month!",
            datetime(2025, 10, 1, 0, 0, 0), datetime(2025, 10, 2, 0, 0, 0),
            "relative monthly", 1, [True, True, True, True, True, False, False],
            until=datetime(2050, 11, 2), all_day=True,
        )
        server_mac.bridge = FakeBridge(raw="", recurring_raw=reminder)

        # A random Thursday that is not the 1st of a month.
        result = asyncio.run(server_mac.list_events("2026-08-27", "2026-08-27", 20))
        self.assertEqual(json.loads(result), [])


if __name__ == "__main__":
    unittest.main()
