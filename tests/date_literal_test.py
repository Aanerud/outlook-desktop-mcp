"""
Unit test for _outlook_date_literal / _outlook_dasl_date_literal — locale-safe
date literals for Outlook Restrict()/DASL filters.

Pure unit test: no Outlook, no pywin32, runs on any platform. The Windows-only
locale formatter (_format_locale_datetime) is replaced with a fake that
emulates a day-first (DD/MM/YYYY) regional setting, so the test checks the
literal-building logic without a Windows machine.

Regression for the locale bug where a hard-coded strftime('%m/%d/%Y %H:%M')
literal such as '09/07/2026 20:15' was parsed by Outlook as 9 Jul instead of
7 Sep on non-US locales, silently returning wrong or empty filter results
(issue #30).

Run: .venv\\Scripts\\python tests\\date_literal_test.py
"""
import os
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server  # noqa: E402

# Record every datetime the literal builders hand to the locale formatter, and
# return a deterministic day-first rendering instead of calling Win32.
_formatted = []


def _fake_format_locale_datetime(dt):
    _formatted.append(dt)
    return f"{dt.day:02d}/{dt.month:02d}/{dt.year:04d} {dt.hour:02d}:{dt.minute:02d}"


server._format_locale_datetime = _fake_format_locale_datetime

_outlook_date_literal = server._outlook_date_literal
_outlook_dasl_date_literal = server._outlook_dasl_date_literal


def test_jet_literal_passes_local_wall_clock_through_unchanged():
    """Jet filters compare in local time: the datetime must reach the locale
    formatter with its wall-clock fields untouched (no tz shift)."""
    _formatted.clear()
    out = _outlook_date_literal(datetime(2026, 9, 7, 20, 15))
    assert out == "07/09/2026 20:15", out
    assert _formatted == [datetime(2026, 9, 7, 20, 15)], _formatted


def test_literal_ordering_comes_from_the_locale_not_the_code():
    """The builder imposes no day/month order of its own — the fake locale is
    day-first, so 7 September stays day-first and never renders as 9 July."""
    out = _outlook_date_literal(datetime(2026, 9, 7, 20, 15))
    assert out.startswith("07/09/2026")
    assert "07" == out.split("/")[0]


def test_no_month_name_is_spliced_into_the_literal():
    """Regression: the earlier fix spliced a fixed English abbreviation
    ('07-Sep-2026') into the literal, which is itself locale-dependent. The
    builder must add no month name of its own."""
    english_abbr = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    for month in range(1, 13):
        out = _outlook_date_literal(datetime(2026, month, 15, 12, 0))
        for abbr in english_abbr:
            assert abbr not in out, (month, out)


def test_dasl_literal_is_converted_to_utc():
    """DASL filters compare in UTC. An aware datetime at +02:00 must reach the
    formatter shifted back to UTC (18:15), while the Jet path keeps 20:15."""
    _formatted.clear()
    aware = datetime(2026, 9, 7, 20, 15, tzinfo=timezone(timedelta(hours=2)))

    assert _outlook_dasl_date_literal(aware) == "07/09/2026 18:15"
    assert _formatted[-1] == datetime(2026, 9, 7, 18, 15)

    assert _outlook_date_literal(aware.replace(tzinfo=None)) == "07/09/2026 20:15"


def test_dasl_literal_leaves_utc_input_unchanged():
    """A datetime already in UTC is not shifted again."""
    _formatted.clear()
    aware_utc = datetime(2026, 9, 7, 20, 15, tzinfo=timezone.utc)
    assert _outlook_dasl_date_literal(aware_utc) == "07/09/2026 20:15"
    assert _formatted[-1] == datetime(2026, 9, 7, 20, 15)


if __name__ == "__main__":
    test_jet_literal_passes_local_wall_clock_through_unchanged()
    test_literal_ordering_comes_from_the_locale_not_the_code()
    test_no_month_name_is_spliced_into_the_literal()
    test_dasl_literal_is_converted_to_utc()
    test_dasl_literal_leaves_utc_input_unchanged()
    print("OK")
