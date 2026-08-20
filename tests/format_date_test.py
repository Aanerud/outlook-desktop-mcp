"""
Unit test for format_date — locale-independent AppleScript date construction.

Pure unit test: no Outlook, no osascript, runs on any platform.

Regression for the locale bug where ``date "2026-03-22 14:00:00"`` was
mis-parsed by AppleScript on non-US locales (e.g. es_ES -> 16 Sep 2027),
scrambling the date while preserving only the time. The fix builds the date
from numeric components via ``run script`` instead of a string literal.
"""
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.utils.applescript_helpers import format_date  # noqa: E402


def test_format_date_uses_numeric_components_not_string_literal():
    expr = format_date(datetime(2026, 3, 22, 14, 5, 9))

    # Must NOT emit a locale-parsed string-literal date.
    assert 'date "2026' not in expr

    # Must build the date from numeric components inside run script.
    assert expr.startswith("(run script ")
    for fragment in (
        "set year of d to 2026",
        "set month of d to 3",
        "set day of d to 22",
        "set hours of d to 14",
        "set minutes of d to 5",
        "set seconds of d to 9",
    ):
        assert fragment in expr, fragment

    # Day is reset to 1 before year/month to avoid end-of-month rollover.
    assert expr.index("set day of d to 1") < expr.index("set month of d to 3")


def test_format_date_is_a_single_inline_expression():
    # No raw newlines: the result is interpolated inline into larger scripts,
    # so newlines must be escaped (\\n) for AppleScript's string literal.
    expr = format_date(datetime(2026, 12, 31, 23, 59, 59))
    assert "\n" not in expr
    assert expr.count("(run script ") == 1


if __name__ == "__main__":
    test_format_date_uses_numeric_components_not_string_literal()
    test_format_date_is_a_single_inline_expression()
    print("OK")
