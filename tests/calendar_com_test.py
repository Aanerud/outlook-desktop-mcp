"""
Outlook Desktop MCP - Calendar COM Validation
===============================================
Standalone script that validates all calendar COM operations before
building the MCP layer. Requires Classic Outlook (Desktop) to be running.

Run: .venv\\Scripts\\python tests\\calendar_com_test.py
"""
import os
import sys
import time
from datetime import datetime, timedelta

# Ensure our package is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.utils.dates import restrict_date


def log(msg):
    print(msg, file=sys.stderr, flush=True)


# Will hold EntryID of the test appointment we create, so later tests can use it
_created_entry_id = None


def test_connect():
    """Test 0: Connect to running Outlook via COM."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    log(f"  Connected to store: {namespace.DefaultStore.DisplayName}")
    return outlook, namespace


def test_list_upcoming_events(namespace, days=7):
    """Test 1: List upcoming events for the next N days, handling recurrences."""
    calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
    items = calendar.Items

    # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    start = datetime.now()
    end = start + timedelta(days=days)
    restrict = (
        f"[Start] >= '{restrict_date(start)}' "
        f"AND [Start] <= '{restrict_date(end)}'"
    )
    filtered = items.Restrict(restrict)

    count = 0
    for item in filtered:
        count += 1
        log(f"  [{count}] {item.Subject}")
        log(f"       {item.Start} - {item.End}")
        log(f"       Location: {item.Location or '(none)'}")
        log(f"       Organizer: {item.Organizer}")
        if count >= 10:
            log(f"  ... (showing first 10)")
            break

    log(f"  Total events in next {days} days: {count}+")
    return count


def test_read_event_details(namespace):
    """Test 2: Read detailed properties from the first upcoming event."""
    calendar = namespace.GetDefaultFolder(9)
    items = calendar.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    start = datetime.now()
    end = start + timedelta(days=30)
    restrict = (
        f"[Start] >= '{restrict_date(start)}' "
        f"AND [Start] <= '{restrict_date(end)}'"
    )
    filtered = items.Restrict(restrict)

    item = None
    for candidate in filtered:
        item = candidate
        break

    if item is None:
        log("  SKIP: No upcoming events found")
        return

    log(f"  Subject: {item.Subject}")
    log(f"  Start: {item.Start}")
    log(f"  End: {item.End}")
    log(f"  Duration: {item.Duration} min")
    log(f"  Location: {item.Location or '(none)'}")
    log(f"  Organizer: {item.Organizer}")
    log(f"  AllDay: {item.AllDayEvent}")
    log(f"  IsRecurring: {item.IsRecurring}")
    log(f"  BusyStatus: {item.BusyStatus}")
    log(f"  MeetingStatus: {item.MeetingStatus}")
    log(f"  RequiredAttendees: {item.RequiredAttendees or '(none)'}")
    log(f"  OptionalAttendees: {item.OptionalAttendees or '(none)'}")
    log(f"  ReminderSet: {item.ReminderSet}")
    log(f"  EntryID: {item.EntryID[:40]}...")
    log(f"  Body: {(item.Body or '')[:100]}...")


def test_create_appointment(outlook):
    """Test 3: Create a personal appointment (no attendees)."""
    global _created_entry_id

    appt = outlook.CreateItem(1)  # olAppointmentItem
    start = datetime.now() + timedelta(days=3, hours=2)
    end = start + timedelta(hours=1)

    appt.Subject = "Outlook Desktop MCP - Calendar COM Test"
    appt.Start = start.strftime("%Y-%m-%d %H:%M")
    appt.End = end.strftime("%Y-%m-%d %H:%M")
    appt.Location = "Test Room"
    appt.Body = "Automated test appointment from calendar COM validation."
    appt.ReminderSet = False
    appt.Save()

    _created_entry_id = appt.EntryID
    log(f"  Created appointment: '{appt.Subject}'")
    log(f"  Start: {appt.Start}")
    log(f"  EntryID: {appt.EntryID[:40]}...")


def test_create_meeting(outlook):
    """Test 4: Create a meeting with an attendee and send the invite."""
    appt = outlook.CreateItem(1)
    start = datetime.now() + timedelta(days=4, hours=3)
    end = start + timedelta(minutes=30)

    appt.Subject = "Outlook Desktop MCP - Calendar Meeting Test"
    appt.Start = start.strftime("%Y-%m-%d %H:%M")
    appt.End = end.strftime("%Y-%m-%d %H:%M")
    appt.Location = "Virtual"
    appt.Body = "Automated meeting test from calendar COM validation."
    appt.MeetingStatus = 1  # olMeeting

    recipient = appt.Recipients.Add("user@example.com")
    recipient.Type = 1  # olRequired
    appt.Recipients.ResolveAll()

    appt.Send()
    log(f"  Meeting sent: '{appt.Subject}' to user@example.com")


def test_update_event(namespace):
    """Test 5: Modify the appointment created in Test 3."""
    global _created_entry_id
    if not _created_entry_id:
        log("  SKIP: No appointment was created in Test 3")
        return

    item = namespace.GetItemFromID(_created_entry_id)
    old_subject = item.Subject
    item.Subject = "Outlook Desktop MCP - Calendar COM Test (UPDATED)"
    item.Location = "Updated Room"
    item.Save()

    # Re-read to verify
    item = namespace.GetItemFromID(_created_entry_id)
    log(f"  Updated: '{old_subject}' -> '{item.Subject}'")
    log(f"  Location: {item.Location}")


def test_delete_event(namespace):
    """Test 6: Delete the appointment created in Test 3."""
    global _created_entry_id
    if not _created_entry_id:
        log("  SKIP: No appointment was created in Test 3")
        return

    item = namespace.GetItemFromID(_created_entry_id)
    subject = item.Subject
    item.Delete()
    _created_entry_id = None
    log(f"  Deleted: '{subject}'")


def test_search_events(namespace, keyword="MCP"):
    """Test 7: Search calendar events by subject keyword.

    Cannot mix DASL @SQL= with regular [Start] property syntax in one
    Restrict call. Use date range filter first, then match subject in Python.
    """
    calendar = namespace.GetDefaultFolder(9)
    items = calendar.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    start = datetime.now() - timedelta(days=30)
    end = datetime.now() + timedelta(days=30)
    restrict = (
        f"[Start] >= '{restrict_date(start)}' "
        f"AND [Start] <= '{restrict_date(end)}'"
    )
    filtered = items.Restrict(restrict)

    keyword_lower = keyword.lower()
    count = 0
    for item in filtered:
        if keyword_lower in (item.Subject or "").lower():
            count += 1
            log(f"  [{count}] {item.Subject} ({item.Start})")
            if count >= 5:
                break

    log(f"  Found {count} events matching '{keyword}'")


def main():
    log("=" * 60)
    log("Outlook Desktop MCP - Calendar COM Validation")
    log("=" * 60)
    log("")

    log("--- Test 0: Connect to Outlook COM ---")
    try:
        outlook, namespace = test_connect()
        log("  PASS")
    except Exception as e:
        log(f"  FAIL: {e}")
        log("\nIs Outlook Desktop (Classic) running?")
        sys.exit(1)

    tests = [
        ("List Upcoming Events (7 days)", lambda: test_list_upcoming_events(namespace)),
        ("Read Event Details", lambda: test_read_event_details(namespace)),
        ("Create Appointment", lambda: test_create_appointment(outlook)),
        ("Create Meeting (send invite)", lambda: test_create_meeting(outlook)),
        ("Update Event", lambda: test_update_event(namespace)),
        ("Delete Event", lambda: test_delete_event(namespace)),
        ("Search Events", lambda: test_search_events(namespace)),
    ]

    passed = 1  # Test 0 already passed
    total = len(tests) + 1

    for name, fn in tests:
        log(f"\n--- Test {passed + 1}: {name} ---")
        try:
            fn()
            passed += 1
            log("  PASS")
        except Exception as e:
            log(f"  FAIL: {e}")

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)

    import pythoncom
    pythoncom.CoUninitialize()
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
