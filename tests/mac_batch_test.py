"""
Outlook Desktop MCP - macOS Batch AppleScript Test
===================================================
Unit tests for the batched (fast) AppleScript generation in server_mac.
Runs without Outlook: the AppleScript bridge is replaced with a fake.

Run: python tests/mac_batch_test.py
"""
import sys
import os
import json
import asyncio

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


def log(msg):
    print(msg, file=sys.stderr, flush=True)


class FakeBridge:
    """Captures scripts and returns canned output."""

    def __init__(self, output="", fail_first_with=None):
        self.scripts = []
        self.output = output
        self.fail_first_with = fail_first_with

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        if self.fail_first_with and len(self.scripts) == 1:
            raise RuntimeError(self.fail_first_with)
        return self.output

    async def run_lines(self, script, timeout=None):
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]


def email_record(mid="101", subject="Hello", sender="a@b.com",
                 sender_name="Alice", time="2026-08-27 10:00:00",
                 is_read="false", att="2"):
    return DELIM.join([mid, subject, sender, sender_name, time, is_read, att])


passed = 0
total = 0


def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        log(f"  PASS: {name}")
    else:
        log(f"  FAIL: {name} {detail}")


def test_list_emails_uses_batch_script():
    log("--- list_emails generates batch script, parses output ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=40)))

    script = fake.scripts[0]
    check("single osascript call", len(fake.scripts) == 1)
    check("batch id fetch", "id of messages 1 thru maxCount" in script, script[:200])
    check("batch subject fetch", "subject of messages 1 thru maxCount" in script)
    check("no per-message subject fetch in loop", "set msubject to subject of m\n" not in script)
    check("no full-folder materialization", "set allMsgs to messages of folderRef" not in script)
    check("one email parsed", len(result) == 1, str(result))
    if result:
        check("subject parsed", result[0]["subject"] == "Hello")
        check("unread derived from is_read", result[0]["unread"] is True)
        check("attachment count parsed", result[0]["attachment_count"] == 2)


def test_list_emails_unread_uses_whose_filter():
    log("--- list_emails unread_only filters server-side ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM)
    server_mac.bridge = fake

    asyncio.run(server_mac.list_emails(folder="inbox", count=10, unread_only=True))
    script = fake.scripts[0]
    check("whose filter present", "whose is read is false" in script)
    check("batch fetch on filtered set", "id of (every message of folderRef whose is read is false)" in script)


def test_list_emails_falls_back_to_legacy():
    log("--- list_emails falls back to legacy loop on AppleScript error ---")
    fake = FakeBridge(output=email_record() + RECORD_DELIM,
                      fail_first_with="AppleScript error: can't batch")
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=5)))
    check("two attempts made", len(fake.scripts) == 2, f"{len(fake.scripts)} scripts")
    if len(fake.scripts) == 2:
        check("legacy loop used second", "set allMsgs to messages of folderRef" in fake.scripts[1])
    check("result parsed after fallback", len(result) == 1 and result[0]["subject"] == "Hello")


def test_list_emails_timeout_not_retried():
    log("--- list_emails does not retry after a timeout ---")
    fake = FakeBridge(fail_first_with="AppleScript timed out after 30s")
    server_mac.bridge = fake

    result = asyncio.run(server_mac.list_emails(folder="inbox", count=5))
    check("only one attempt after timeout", len(fake.scripts) == 1, f"{len(fake.scripts)} scripts")
    check("error surfaced", "timed out" in result)


def test_search_emails_batches_filtered_set():
    log("--- search_emails batches over whose-filtered set ---")
    fake = FakeBridge(output=email_record(subject="budget") + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.search_emails(query="budget", count=10)))
    script = fake.scripts[0]
    check("whose subject filter", 'whose subject contains "budget"' in script)
    check("batch fetch present", "id of (every message of folderRef whose subject contains" in script)
    check("no per-message property loop", "set msubject to subject of m\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "budget")


def event_record(eid="7", subject="Standup", start="2026-08-27T09:00:00",
                 end="2026-08-27T09:15:00", loc="Room 1", org="Bob", allday="false"):
    return DELIM.join([eid, subject, start, end, loc, org, allday])


def test_list_events_batches():
    log("--- list_events filters by date server-side and batches fetch ---")
    fake = FakeBridge(output=event_record() + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=20)))
    script = fake.scripts[0]
    check("whose date filter",
          "whose start time ≥ rangeStart and start time ≤ rangeEnd" in script)
    check("locale-safe date construction", "set year of rangeStart to 2026" in script
          and "set month of rangeStart to 8" in script)
    check("iso date output", "«class isot»" in script)
    check("no per-event property loop", "set esubject to subject of e\n" not in script)
    check("no full calendar materialization", "set evts to calendar events\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "Standup", str(result))


def test_list_events_filters_and_sorts_in_python():
    log("--- list_events drops out-of-range events and sorts by start ---")
    recs = RECORD_DELIM.join([
        event_record(eid="3", subject="Late", start="2026-08-27T15:00:00", end="2026-08-27T16:00:00"),
        event_record(eid="1", subject="Old holiday", start="2012-01-01T00:00:00", end="2012-01-02T00:00:00"),
        event_record(eid="2", subject="Early", start="2026-08-27T09:00:00", end="2026-08-27T09:15:00"),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=20)))
    check("out-of-range event dropped", all(e["subject"] != "Old holiday" for e in result), str(result))
    check("two in-range events kept", len(result) == 2, str(result))
    check("sorted by start time", [e["subject"] for e in result] == ["Early", "Late"], str(result))


def test_list_events_truncates_to_count():
    log("--- list_events returns at most count events after sorting ---")
    recs = RECORD_DELIM.join([
        event_record(eid=str(i), subject=f"E{i}", start=f"2026-08-27T{9 + i:02d}:00:00",
                     end=f"2026-08-27T{10 + i:02d}:00:00")
        for i in range(3)
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_events(
        start_date="2026-08-27", end_date="2026-08-28", count=2)))
    check("truncated to count", len(result) == 2, str(result))
    check("earliest kept", [e["subject"] for e in result] == ["E0", "E1"], str(result))


def test_list_folders_batches_names():
    log("--- list_folders enumerates per account with batched fetches ---")
    recs = RECORD_DELIM.join([
        DELIM.join(["edd@example.com", "Inbox", "1200", "4"]),
        DELIM.join(["edd@example.com", "Inbox/kapptivate", "13413", "0"]),
        DELIM.join(["On My Computer", "Inbox", "0", "0"]),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_folders()))
    script = fake.scripts[0]
    check("enumerates exchange accounts", "exchange accounts" in script)
    check("enumerates local account", "on my computer" in script)
    check("batch name fetch", "name of mail folders of" in script)
    check("batch unread fetch", "unread count of mail folders of" in script)
    check("nameless containers skipped in script", "is not missing value" in script)
    check("no per-folder name fetch", "set fname to name of f\n" not in script)
    check("account tagged", result and result[0].get("account") == "edd@example.com", str(result))
    check("subfolder path kept", any(f["name"] == "Inbox/kapptivate" for f in result), str(result))
    check("local account tagged", any(f["account"] == "On My Computer" for f in result), str(result))
    check("counts parsed", result and result[0]["item_count"] == 1200 and result[0]["unread_count"] == 4)


def test_list_folders_depth_and_legacy():
    log("--- list_folders honors max_depth and drops nameless legacy rows ---")
    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=1))
    check("depth 1 skips subfolder pass", "mail folders of (item i1 of refs1)" not in fake.scripts[0])

    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=2))
    check("depth 2 includes subfolder pass", "mail folders of (item i1 of refs1)" in fake.scripts[0])

    fake = FakeBridge(output="")
    server_mac.bridge = fake
    asyncio.run(server_mac.list_folders(max_depth=3))
    check("depth 3 recurses further", "mail folders of (item i2 of refs2)" in fake.scripts[0])
    check("subfolder path built", 'path1 & "/" & n2' in fake.scripts[0])

    # Legacy 3-field records (flat fallback script) still parse; nameless dropped.
    recs = RECORD_DELIM.join([
        DELIM.join(["Inbox", "10", "2"]),
        DELIM.join(["missing value", "0", "0"]),
    ]) + RECORD_DELIM
    fake = FakeBridge(output=recs)
    server_mac.bridge = fake
    result = json.loads(asyncio.run(server_mac.list_folders()))
    check("legacy record parsed", any(f["name"] == "Inbox" and f["item_count"] == 10 for f in result), str(result))
    check("nameless row dropped", all(f["name"] != "missing value" for f in result), str(result))


def test_list_tasks_batches():
    log("--- list_tasks batches task fetch ---")
    rec = DELIM.join(["55", "Pay invoice", "2026-09-01 00:00:00", "not completed", "priority normal"])
    fake = FakeBridge(output=rec + RECORD_DELIM)
    server_mac.bridge = fake

    result = json.loads(asyncio.run(server_mac.list_tasks(count=20)))
    script = fake.scripts[0]
    check("batch id fetch on filtered tasks", "id of (every task whose todo flag is not completed)" in script)
    check("no per-task property loop", "set tname to name of t\n" not in script)
    check("result parsed", len(result) == 1 and result[0]["subject"] == "Pay invoice")


def test_script_timeout_env_override():
    log("--- SCRIPT_TIMEOUT honors OUTLOOK_MCP_SCRIPT_TIMEOUT env var ---")
    import importlib
    from outlook_desktop_mcp import applescript_bridge

    os.environ["OUTLOOK_MCP_SCRIPT_TIMEOUT"] = "45"
    try:
        importlib.reload(applescript_bridge)
        check("env override applied", applescript_bridge.SCRIPT_TIMEOUT == 45.0,
              str(applescript_bridge.SCRIPT_TIMEOUT))
    finally:
        del os.environ["OUTLOOK_MCP_SCRIPT_TIMEOUT"]
        importlib.reload(applescript_bridge)
    check("default raised above 30s", applescript_bridge.SCRIPT_TIMEOUT >= 60,
          str(applescript_bridge.SCRIPT_TIMEOUT))


def main():
    test_list_emails_uses_batch_script()
    test_list_emails_unread_uses_whose_filter()
    test_list_emails_falls_back_to_legacy()
    test_list_emails_timeout_not_retried()
    test_search_emails_batches_filtered_set()
    test_list_events_batches()
    test_list_events_filters_and_sorts_in_python()
    test_list_events_truncates_to_count()
    test_list_folders_batches_names()
    test_list_folders_depth_and_legacy()
    test_list_tasks_batches()
    test_script_timeout_env_override()

    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
