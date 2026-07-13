from datetime import datetime

from outlook_desktop_mcp.server import _collect_calendar_items


class FakeItem:
    def __init__(self, start: datetime):
        self.Start = start


def test_collect_calendar_items_inclusive_range_across_month_boundary():
    items = [
        FakeItem(datetime(2026, 7, 21, 9, 0)),
        FakeItem(datetime(2026, 7, 27, 9, 0)),
        FakeItem(datetime(2026, 8, 1, 9, 0)),
        FakeItem(datetime(2026, 8, 7, 9, 0)),
        FakeItem(datetime(2026, 9, 1, 9, 0)),
    ]

    results = _collect_calendar_items(
        items,
        datetime(2026, 8, 1, 0, 0),
        datetime(2026, 8, 7, 23, 59),
        20,
    )

    assert [item.Start for item in results] == [
        datetime(2026, 8, 1, 9, 0),
        datetime(2026, 8, 7, 9, 0),
    ]


def test_collect_calendar_items_stops_after_count():
    items = [
        FakeItem(datetime(2026, 8, 1, 9, 0)),
        FakeItem(datetime(2026, 8, 2, 9, 0)),
        FakeItem(datetime(2026, 8, 3, 9, 0)),
    ]

    results = _collect_calendar_items(
        items,
        datetime(2026, 8, 1, 0, 0),
        datetime(2026, 8, 31, 23, 59),
        2,
    )

    assert [item.Start for item in results] == [
        datetime(2026, 8, 1, 9, 0),
        datetime(2026, 8, 2, 9, 0),
    ]
