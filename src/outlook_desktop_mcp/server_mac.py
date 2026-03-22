"""
Outlook Desktop MCP Server — macOS
====================================
Exposes Microsoft Outlook for Mac as an MCP server over stdio.
Uses AppleScript automation via osascript — no Microsoft Graph, no Entra app.
Just run this on macOS with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp (auto-detected on macOS)
"""
import sys
import json
import logging
import os
import re

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.applescript_bridge import AppleScriptBridge
from outlook_desktop_mcp.utils.applescript_helpers import (
    escape,
    format_date,
    parse_date,
    resolve_folder_ref,
    DELIM,
    RECORD_DELIM,
)
from datetime import datetime, timedelta

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")

# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook on macOS "
        "via AppleScript automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders, reply to emails, manage calendar events, and "
        "manage tasks.\n\n"
        "All operations use the locally running Outlook app — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Microsoft Outlook for Mac must be running.\n\n"
        "NOTE: entry_id values on macOS are numeric IDs (not hex strings like "
        "on Windows). They identify items within their folder context.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "search events\n"
        "- Tasks: create, list, complete, delete to-do items\n"
        "- Attachments: list and save attachments\n"
        "- Folders: list folder hierarchy"
    ),
)

bridge = AppleScriptBridge()


# --- Helper: truncate long text ---

def _truncate(text: str, max_length: int = 5000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def _clean(value: str) -> str:
    """Replace AppleScript's 'missing value' with empty string."""
    v = value.strip()
    return "" if v == "missing value" else v


# --- UI Scraping for New Outlook for Mac ---
# New Outlook for Mac stores Exchange/M365 mailbox data in the cloud and
# does NOT expose it through the AppleScript `inbox` keyword (which only
# reaches the empty local "On My Computer" inbox). The only way to access
# Exchange messages is via macOS UI scripting (System Events), reading
# the message list table visible in the Outlook window.


_UI_MESSAGE_LIST_PATH = (
    'tell application "System Events"\n'
    '    tell process "Microsoft Outlook"\n'
    '        tell window 1\n'
    '            tell splitter group 1\n'
    '                tell splitter group 1\n'
    '                    tell splitter group 1\n'
    '                        tell group 1\n'
    '                            tell scroll area 1\n'
    '                                tell table 1\n'
)

_UI_MESSAGE_LIST_END = (
    '                                end tell\n'
    '                            end tell\n'
    '                        end tell\n'
    '                    end tell\n'
    '                end tell\n'
    '            end tell\n'
    '        end tell\n'
    '    end tell\n'
    'end tell'
)


async def _ui_list_messages(bridge_obj, count: int = 10) -> list[dict]:
    """Read visible inbox messages via UI scripting (System Events).

    This is the fallback for New Outlook for Mac where AppleScript's inbox
    keyword only sees the empty local mailbox.
    """
    script = (
        _UI_MESSAGE_LIST_PATH +
        f'                                    set rowList to rows\n'
        f'                                    set rowCount to count of rowList\n'
        f'                                    set maxRows to rowCount\n'
        f'                                    if maxRows > {count} then set maxRows to {count}\n'
        f'                                    set output to ""\n'
        f'                                    repeat with i from 1 to maxRows\n'
        f'                                        set r to row i\n'
        f'                                        try\n'
        f'                                            set cellDesc to description of UI element 1 of r\n'
        f'                                            set output to output & cellDesc & "{RECORD_DELIM}"\n'
        f'                                        end try\n'
        f'                                    end repeat\n'
        f'                                    return output\n' +
        _UI_MESSAGE_LIST_END
    )

    raw = await bridge_obj.run(script)
    if not raw:
        return []

    results = []
    for idx, record in enumerate(raw.split(RECORD_DELIM), start=1):
        record = record.strip()
        if not record:
            continue
        # Cell description format uses `,` + 4+ spaces as major field
        # separators, while in-content commas have 0-1 trailing spaces.
        # Structure: [Ulest,]    SENDER, SUBJECT,     TIME,    [FLAGS,]
        fields = [f.strip() for f in re.split(r",\s{4,}", record)]

        is_unread = False
        has_attachment = False
        # Strip status flags from the fields
        cleaned = []
        for f in fields:
            if not f:
                continue
            if f == "Ulest":
                is_unread = True
                continue
            if "Har filer" in f:
                has_attachment = True
                continue
            if f == "Kategorisert" or f.startswith("Merket som"):
                continue
            cleaned.append(f)

        # cleaned is typically: [SENDER_AND_SUBJECT, TIME]
        # or [SENDER_AND_SUBJECT, TIME, extra...]
        # SENDER_AND_SUBJECT is: "Sender, Subject" (comma + 1 space)
        sender_subject = cleaned[0] if cleaned else ""
        time_str = cleaned[1] if len(cleaned) > 1 else ""
        # Strip trailing comma from time
        time_str = time_str.rstrip(",").strip()

        # Split sender from subject on first ", " (comma + single space)
        # But skip thread-count prefixes like "2 meldinger, "
        ss = sender_subject
        # Remove thread count prefix
        ss = re.sub(r"^\d+\s+meldinger,\s*", "", ss)
        # Split on first ", " to get sender and subject
        comma_pos = ss.find(", ")
        if comma_pos > 0:
            sender = ss[:comma_pos].strip()
            subject = ss[comma_pos + 2:].strip()
        else:
            sender = ""
            subject = ss.strip()

        results.append({
            "entry_id": f"ui-{idx}",
            "subject": subject or "(could not parse subject)",
            "sender": "",
            "sender_name": sender,
            "received_time": time_str,
            "unread": is_unread,
            "has_attachments": has_attachment,
            "attachment_count": 1 if has_attachment else 0,
            "_source": "ui_scraping",
        })

    return results


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    # Build recipient lines
    def _recipient_lines(addresses: str, kind: str) -> str:
        lines = ""
        for addr in addresses.split(";"):
            addr = addr.strip()
            if addr:
                lines += f'make new {kind} at newMsg with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
        return lines

    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""
    bcc_lines = _recipient_lines(bcc, "bcc recipient") if bcc else ""

    content_prop = f'html content:"{escape(html_body)}"' if html_body else f'content:"{escape(body)}"'

    script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}
    send newMsg
end tell'''

    try:
        await bridge.run(script)
        return f"Email sent: '{subject}' to {to}"
    except Exception as e:
        return f"Error sending email: {e}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", or any custom folder name visible in list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.

    Returns:
        JSON array of email summary objects.
    """
    folder_ref = resolve_folder_ref(folder)
    unread_filter = ' whose is read is false' if unread_only else ''

    script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set allMsgs to messages of folderRef{unread_filter}
    set msgCount to count of allMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of allMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)

        results = []
        if raw:
            for record in raw.split(RECORD_DELIM):
                record = record.strip()
                if not record:
                    continue
                parts = record.split(DELIM)
                if len(parts) < 7:
                    continue
                att_count = int(parts[6]) if parts[6].strip().isdigit() else 0
                results.append({
                    "entry_id": parts[0].strip(),
                    "subject": parts[1].strip() or "(no subject)",
                    "sender": parts[2].strip(),
                    "sender_name": parts[3].strip(),
                    "received_time": _clean(parts[4]),
                    "unread": parts[5].strip().lower() != "true",  # is_read -> unread
                    "has_attachments": att_count > 0,
                    "attachment_count": att_count,
                })

        # Fallback: New Outlook for Mac keeps Exchange messages outside the
        # AppleScript-visible mailbox. If the standard query returned nothing
        # for the inbox, try reading the visible message list via UI scripting.
        if not results and folder.lower().strip() in ("inbox", "innboks"):
            try:
                results = await _ui_list_messages(bridge, count)
            except Exception:
                pass  # UI scraping failed — return empty list

        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing emails: {e}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The numeric ID of the email. Most reliable way to identify
            a specific email. Get this from list_emails or search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    if entry_id:
        folder_ref = resolve_folder_ref(folder)
        script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m as string
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
    elif subject_search:
        folder_ref = resolve_folder_ref(folder)
        safe_query = escape(subject_search)
        script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    if (count of matchMsgs) = 0 then return "NOT_FOUND"
    set m to item 1 of matchMsgs
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m as string
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{DELIM}" & mto & "{DELIM}" & mcc & "{DELIM}" & mbody
end tell'''
    else:
        return json.dumps({"error": "Provide either entry_id or subject_search"})

    try:
        raw = await bridge.run(script)
        if raw == "NOT_FOUND":
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        parts = raw.split(DELIM, 9)  # max 10 parts
        if len(parts) < 10:
            return json.dumps({"error": "Failed to parse email data"})

        att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "sender": parts[2].strip(),
            "sender_name": parts[3].strip(),
            "received_time": _clean(parts[4]),
            "unread": parts[5].strip().lower() != "true",
            "has_attachments": att_count > 0,
            "attachment_count": att_count,
            "to": parts[7].strip(),
            "cc": parts[8].strip(),
            "body": _truncate(_clean(parts[9])),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading email: {e}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str) -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to true
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as read: '{subject}'"
    except Exception as e:
        return f"Error marking email as read: {e}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str) -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to false
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as unread: '{subject}'"
    except Exception as e:
        return f"Error marking email as unread: {e}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid.

    Args:
        entry_id: The numeric ID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "inbox", "sent", "deleted"/"trash",
            "drafts", "junk"/"spam", or any custom folder name.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    dest_ref = resolve_folder_ref(target_folder)
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    move m to {dest_ref}
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Moved '{subject}' to {target_folder}"
    except Exception as e:
        return f"Error moving email: {e}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The numeric ID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    reply_cmd = "reply all to" if reply_all else "reply to"
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    set replyMsg to {reply_cmd} m
    set content of replyMsg to "{escape(body)}" & return & return & content of replyMsg
    send replyMsg
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Reply sent to '{subject}' (reply_all={reply_all})"
    except Exception as e:
        return f"Error replying to email: {e}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(max_depth: int = 2) -> str:
    """List all mail folders in the user's Outlook mailbox.

    Returns a JSON array showing the folder hierarchy with item counts.
    Use this to discover folder names for other tools (list_emails,
    move_email, search_emails).

    Args:
        max_depth: How many levels deep to recurse into subfolders.
            Default 2. Set to 1 for top-level only.

    Returns:
        JSON array of folder objects with name, item_count, and unread_count.
    """
    script = f'''tell application "Microsoft Outlook"
    set allFolders to mail folders
    set output to ""
    repeat with f in allFolders
        set fname to name of f
        set fcount to count of messages of f
        set funread to unread count of f
        set output to output & fname & "{DELIM}" & (fcount as text) & "{DELIM}" & (funread as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append({
                "name": parts[0].strip(),
                "item_count": int(parts[1].strip()) if parts[1].strip().isdigit() else 0,
                "unread_count": int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing folders: {e}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects using Outlook's AppleScript filtering.
    Results include entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    folder_ref = resolve_folder_ref(folder)
    safe_query = escape(query)

    script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    set msgCount to count of matchMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of matchMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m as string
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & (mid as text) & "{DELIM}" & msubject & "{DELIM}" & msender & "{DELIM}" & msenderName & "{DELIM}" & mtime & "{DELIM}" & (misread as text) & "{DELIM}" & (mattcount as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "sender": parts[2].strip(),
                "sender_name": parts[3].strip(),
                "received_time": _clean(parts[4]),
                "unread": parts[5].strip().lower() != "true",
                "has_attachments": att_count > 0,
                "attachment_count": att_count,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error searching emails: {e}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Each summary has entry_id, subject, start, end, duration,
    location, organizer, and attendee info.

    Use entry_id from results with get_event, update_event, or delete_event.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.

    Returns:
        JSON array of event summary objects.
    """
    start = datetime.fromisoformat(start_date) if start_date else datetime.now()
    end = datetime.fromisoformat(end_date) if end_date else start + timedelta(days=7)

    # Fetch more than needed, filter by date in Python since AppleScript
    # whose-clause date filtering can be unreliable in Outlook for Mac.
    fetch_limit = count * 3  # overfetch to account for out-of-range events

    script = f'''tell application "Microsoft Outlook"
    set evts to calendar events
    set evtCount to count of evts
    set maxFetch to {fetch_limit}
    if evtCount < maxFetch then set maxFetch to evtCount
    set output to ""
    repeat with i from 1 to maxFetch
        set e to item i of evts
        set eid to id of e
        set esubject to subject of e
        set estart to start time of e as string
        set eend to end time of e as string
        set elocation to ""
        try
            set elocation to location of e
        end try
        set eorganizer to ""
        try
            set eorganizer to organizer of e
        end try
        set eallday to all day flag of e
        set output to output & (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "start": parts[2].strip(),
                "end": parts[3].strip(),
                "location": _clean(parts[4]),
                "organizer": _clean(parts[5]),
                "all_day": parts[6].strip().lower() == "true",
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing events: {e}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str) -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, and recurrence status.

    Args:
        entry_id: The numeric ID of the event. Get this from list_events
            or search_events results.

    Returns:
        JSON object with full event details.
    """
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set eid to id of e
    set esubject to subject of e
    set estart to start time of e as string
    set eend to end time of e as string
    set elocation to ""
    try
        set elocation to location of e
    end try
    set eorganizer to ""
    try
        set eorganizer to organizer of e
    end try
    set eallday to all day flag of e
    set ebody to ""
    try
        set ebody to plain text content of e
    end try
    set eattendees to ""
    try
        set attList to attendees of e
        repeat with a in attList
            set eattendees to eattendees & address of a & "; "
        end repeat
    end try
    return (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{DELIM}" & ebody & "{DELIM}" & eattendees
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 8)
        if len(parts) < 9:
            return json.dumps({"error": "Failed to parse event data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "start": parts[2].strip(),
            "end": parts[3].strip(),
            "location": _clean(parts[4]),
            "organizer": _clean(parts[5]),
            "all_day": parts[6].strip().lower() == "true",
            "body": _truncate(_clean(parts[7])),
            "attendees": parts[8].strip(),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading event: {e}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date.
        end: End time in ISO 8601 format.
        location: Optional. Event location.
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)

    props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(body)}"'
    if all_day:
        props += ', all day flag:true'

    script = f'''tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    return (id of newEvt as text) & "{DELIM}" & (subject of newEvt) & "{DELIM}" & (start time of newEvt as string) & "{DELIM}" & (end time of newEvt as string)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "start": parts[2].strip() if len(parts) > 2 else start,
            "end": parts[3].strip() if len(parts) > 3 else end,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating event: {e}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and sends meeting requests to all specified
    attendees. Note: invitation delivery depends on Outlook's AppleScript
    meeting support on macOS.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location.
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.

    Returns:
        Confirmation that the meeting was created.
    """
    start_dt = datetime.fromisoformat(start)
    end_dt = datetime.fromisoformat(end)

    props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(body)}"'

    attendee_lines = ""
    for addr in required_attendees.split(";"):
        addr = addr.strip()
        if addr:
            attendee_lines += f'make new required attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
    if optional_attendees:
        for addr in optional_attendees.split(";"):
            addr = addr.strip()
            if addr:
                attendee_lines += f'make new optional attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'

    script = f'''tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    {attendee_lines}
    return (id of newEvt as text)
end tell'''

    try:
        await bridge.run(script)
        return f"Meeting '{subject}' created with attendees: {required_attendees}"
    except Exception as e:
        return f"Error creating meeting: {e}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged.

    Args:
        entry_id: The numeric ID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.

    Returns:
        Confirmation with updated event details, or an error.
    """
    set_lines = ""
    if subject:
        set_lines += f'set subject of e to "{escape(subject)}"\n'
    if start:
        start_dt = datetime.fromisoformat(start)
        set_lines += f'set start time of e to {format_date(start_dt)}\n'
    if end:
        end_dt = datetime.fromisoformat(end)
        set_lines += f'set end time of e to {format_date(end_dt)}\n'
    if location:
        set_lines += f'set location of e to "{escape(location)}"\n'
    if body:
        set_lines += f'set content of e to "{escape(body)}"\n'

    if not set_lines:
        return json.dumps({"error": "No fields to update"})

    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    {set_lines}
    return (id of e as text) & "{DELIM}" & (subject of e) & "{DELIM}" & (start time of e as string) & "{DELIM}" & (end time of e as string) & "{DELIM}" & (location of e)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "updated",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "start": parts[2].strip() if len(parts) > 2 else "",
            "end": parts[3].strip() if len(parts) > 3 else "",
            "location": _clean(parts[4]) if len(parts) > 4 else "",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error updating event: {e}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str) -> str:
    """Delete a calendar event.

    For personal appointments, the event is simply deleted. For meetings,
    the event is removed from your calendar.

    Args:
        entry_id: The numeric ID of the event to delete.

    Returns:
        Confirmation with the event subject, or an error.
    """
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set esubject to subject of e
    delete e
    return esubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Event deleted: '{subject}'"
    except Exception as e:
        return f"Error deleting event: {e}"


# =====================================================================
# TOOL 16: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching event summaries.
    """
    safe_query = escape(query)

    script = f'''tell application "Microsoft Outlook"
    set evts to calendar events whose subject contains "{safe_query}"
    set evtCount to count of evts
    set maxCount to {count}
    if evtCount < maxCount then set maxCount to evtCount
    set output to ""
    repeat with i from 1 to maxCount
        set e to item i of evts
        set eid to id of e
        set esubject to subject of e
        set estart to start time of e as string
        set eend to end time of e as string
        set elocation to ""
        try
            set elocation to location of e
        end try
        set eorganizer to ""
        try
            set eorganizer to organizer of e
        end try
        set eallday to all day flag of e
        set output to output & (eid as text) & "{DELIM}" & esubject & "{DELIM}" & estart & "{DELIM}" & eend & "{DELIM}" & elocation & "{DELIM}" & eorganizer & "{DELIM}" & (eallday as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "start": parts[2].strip(),
                "end": parts[3].strip(),
                "location": _clean(parts[4]),
                "organizer": _clean(parts[5]),
                "all_day": parts[6].strip().lower() == "true",
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error searching events: {e}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries. Each task includes entry_id,
    subject, due_date, and completion status.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.

    Returns:
        JSON array of task summary objects.
    """
    completed_filter = "" if include_completed else " whose todo flag is not completed"

    script = f'''tell application "Microsoft Outlook"
    set taskList to tasks{completed_filter}
    set taskCount to count of taskList
    set maxCount to {count}
    if taskCount < maxCount then set maxCount to taskCount
    set output to ""
    repeat with i from 1 to maxCount
        set t to item i of taskList
        set tid to id of t
        set tname to name of t
        set tdue to ""
        try
            set tdue to due date of t as string
        end try
        set tflag to todo flag of t
        set tpriority to priority of t
        set output to output & (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 5:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "due_date": _clean(parts[2]) or None,
                "complete": parts[3].strip() == "completed",
                "priority": parts[4].strip(),
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing tasks: {e}"


@mcp.tool()
async def get_task(entry_id: str) -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        JSON object with full task details including body.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tid to id of t
    set tname to name of t
    set tdue to ""
    try
        set tdue to due date of t as string
    end try
    set tflag to todo flag of t
    set tpriority to priority of t
    set tbody to ""
    try
        set tbody to plain text content of t
    end try
    set tstartdate to ""
    try
        set tstartdate to start date of t as string
    end try
    return (tid as text) & "{DELIM}" & tname & "{DELIM}" & tdue & "{DELIM}" & (tflag as text) & "{DELIM}" & (tpriority as text) & "{DELIM}" & tbody & "{DELIM}" & tstartdate
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 6)
        if len(parts) < 7:
            return json.dumps({"error": "Failed to parse task data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "due_date": _clean(parts[2]) or None,
            "complete": parts[3].strip() == "completed",
            "priority": parts[4].strip(),
            "body": _truncate(_clean(parts[5])),
            "start_date": _clean(parts[6]) or None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading task: {e}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".

    Returns:
        Confirmation with task subject and entry_id.
    """
    imp_map = {"low": "priority low", "normal": "priority normal", "high": "priority high"}
    imp_val = imp_map.get(importance.lower(), "priority normal")

    props = f'name:"{escape(subject)}", priority:{imp_val}'
    if due_date:
        due_dt = datetime.fromisoformat(due_date)
        props += f', due date:{format_date(due_dt)}'
    if body:
        props += f', content:"{escape(body)}"'

    script = f'''tell application "Microsoft Outlook"
    set newTask to make new task with properties {{{props}}}
    return (id of newTask as text) & "{DELIM}" & (name of newTask)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "due_date": due_date or None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating task: {e}"


@mcp.tool()
async def complete_task(entry_id: str) -> str:
    """Mark a task as complete.

    Sets the task status to complete.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        Confirmation with the task subject.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set todo flag of t to completed
    return name of t
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task completed: '{name}'"
    except Exception as e:
        return f"Error completing task: {e}"


@mcp.tool()
async def delete_task(entry_id: str) -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The numeric ID of the task to delete.

    Returns:
        Confirmation with the task subject.
    """
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tname to name of t
    delete t
    return tname
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task deleted: '{name}'"
    except Exception as e:
        return f"Error deleting task: {e}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str) -> str:
    """List all attachments on an email.

    Args:
        entry_id: The numeric ID of the email to check for attachments.

    Returns:
        JSON array of attachment objects with index and filename.
    """
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    set output to ""
    repeat with i from 1 to attCount
        set a to item i of attList
        set aname to name of a
        set asize to file size of a
        set output to output & (i as text) & "{DELIM}" & aname & "{DELIM}" & (asize as text) & "{RECORD_DELIM}"
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append({
                "index": int(parts[0].strip()) if parts[0].strip().isdigit() else 0,
                "filename": parts[1].strip(),
                "size": int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing attachments: {e}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
) -> str:
    """Save an attachment from an email to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The numeric ID of the email containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1.
            Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    if not save_directory:
        save_directory = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_directory, exist_ok=True)

    # Use POSIX path for AppleScript
    save_dir_posix = save_directory

    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s), requested index {attachment_index}"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to POSIX file "{escape(save_dir_posix)}/{escape("__PLACEHOLDER__")}"
    save a in file ((POSIX path of (POSIX file "{escape(save_dir_posix)}")) & aname)
    return aname
end tell'''

    # Simpler approach: save to known path
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    set a to item {attachment_index} of attList
    set aname to name of a
    set savePath to "{escape(save_dir_posix)}/" & aname
    save a in savePath
    return aname & "{DELIM}" & savePath
end tell'''

    try:
        raw = await bridge.run(script)
        if raw.startswith("ERROR:"):
            return raw

        parts = raw.split(DELIM)
        filename = parts[0].strip() if len(parts) > 0 else "unknown"
        save_path = os.path.join(save_directory, filename)
        result = {
            "status": "saved",
            "filename": filename,
            "path": save_path,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error saving attachment: {e}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    import asyncio

    async def _start():
        logger.info("Starting Outlook Desktop MCP server (macOS)...")
        await bridge.start()
        logger.info("AppleScript bridge ready. Starting MCP stdio transport...")

    asyncio.run(_start())
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()
