"""Date helpers for Outlook COM filters."""
import ctypes
from datetime import datetime, timezone


class _SYSTEMTIME(ctypes.Structure):
    _fields_ = [
        ("wYear", ctypes.c_uint16),
        ("wMonth", ctypes.c_uint16),
        ("wDayOfWeek", ctypes.c_uint16),
        ("wDay", ctypes.c_uint16),
        ("wHour", ctypes.c_uint16),
        ("wMinute", ctypes.c_uint16),
        ("wSecond", ctypes.c_uint16),
        ("wMilliseconds", ctypes.c_uint16),
    ]


_DATE_SHORTDATE = 0x1
_TIME_NOSECONDS = 0x2


def restrict_date(dt: datetime, utc: bool = False) -> str:
    """Format a datetime for Items.Restrict filters.

    Outlook parses date strings in Restrict using the Windows user locale
    (e.g. dd.MM.yyyy on German systems), so a hardcoded US format swaps
    day and month or fails outright. Format with the locale's own short
    date and time patterns instead.

    utc=True converts a naive local datetime to UTC first — required for
    DASL (@SQL=) date properties, which are compared in UTC. Jet-syntax
    properties like [ReceivedTime] and [Start] use local time.
    """
    if utc:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    st = _SYSTEMTIME(dt.year, dt.month, dt.isoweekday() % 7, dt.day,
                     dt.hour, dt.minute, 0, 0)
    kernel32 = ctypes.windll.kernel32
    date_buf = ctypes.create_unicode_buffer(128)
    time_buf = ctypes.create_unicode_buffer(128)
    # None locale name = LOCALE_NAME_USER_DEFAULT
    if not kernel32.GetDateFormatEx(None, _DATE_SHORTDATE, ctypes.byref(st),
                                    None, date_buf, 128, None):
        raise ctypes.WinError()
    if not kernel32.GetTimeFormatEx(None, _TIME_NOSECONDS, ctypes.byref(st),
                                    None, time_buf, 128):
        raise ctypes.WinError()
    return f"{date_buf.value} {time_buf.value}"
