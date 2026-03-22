"""Helpers for building and parsing AppleScript safely."""
import re
from datetime import datetime


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, and other special characters.
    """
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    return text


def format_date(dt: datetime) -> str:
    """Convert a Python datetime to an AppleScript date string.

    Returns a string like: date "Sunday, March 22, 2026 at 2:00:00 PM"
    AppleScript parses dates based on the system locale, so we use a
    locale-friendly format that osascript can interpret.
    """
    return f'date "{dt.strftime("%Y-%m-%d %H:%M:%S")}"'


def parse_date(text: str) -> str:
    """Parse an AppleScript date string to ISO 8601 format.

    AppleScript dates look like: "Sunday, March 22, 2026 at 2:00:00 PM"
    or various locale-specific formats. We attempt several common patterns.
    """
    text = text.strip()
    # Remove day name prefix if present (e.g., "Sunday, ")
    text = re.sub(r"^\w+day,\s*", "", text)
    # Remove " at " between date and time
    text = text.replace(" at ", " ")
    # Try common formats
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",   # March 22, 2026 2:00:00 PM
        "%d. %B %Y %H:%M:%S",       # 22. mars 2026 14:00:00 (Norwegian)
        "%Y-%m-%d %H:%M:%S",        # 2026-03-22 14:00:00
        "%d/%m/%Y %H:%M:%S",        # 22/03/2026 14:00:00
        "%m/%d/%Y %H:%M:%S",        # 03/22/2026 14:00:00
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    # Fallback: return as-is
    return text


# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference.

    Returns an AppleScript expression like 'inbox' or 'mail folder "Archive"'.
    Built-in folders use locale-independent keywords; custom folders use name lookup.
    """
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    # Custom folder — search by name
    return f'mail folder "{escape(folder_name)}"'


# Delimiter used for structured AppleScript output
DELIM = "|||"
RECORD_DELIM = "==="
