"""
Outlook Desktop MCP - macOS email sender/recipient field tests
==============================================================
Guards the fix in server_mac.py for empty sender / sender_name / to / cc
fields on Legacy Outlook for Mac.

Two layers of coverage, both offline (no live Outlook needed):

1. Parse layer — a FakeBridge returns DELIM/RECORD_DELIM-encoded output in
   the shape a *correct* AppleScript run produces (populated address fields),
   and we assert list_emails / read_email surface those fields non-empty.

2. Script layer — we capture the AppleScript that each tool generates and
   assert it uses the corrected object-model access patterns
   (`set snd to sender of m`, `email address of r` -> `address of rea`) and
   NOT the broken inline patterns that returned empty strings.
"""
import asyncio
import json
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


def _summary_record(entry_id, subject, sender, sender_name,
                    mtime="Monday, 3 March 2026 at 09:00:00",
                    is_read="false", att_count="0"):
    """Encode one list_emails row exactly as the fixed AppleScript would."""
    return DELIM.join([
        entry_id, subject, sender, sender_name, mtime, is_read, att_count,
    ])


def _full_record(entry_id, subject, sender, sender_name, to, cc, body,
                 mtime="Monday, 3 March 2026 at 09:00:00",
                 is_read="false", att_count="0"):
    """Encode one read_email row exactly as the fixed AppleScript would."""
    return DELIM.join([
        entry_id, subject, sender, sender_name, mtime, is_read, att_count,
        to, cc, body,
    ])


class FakeBridge:
    """Captures the script it was asked to run and returns canned output."""

    def __init__(self, raw):
        self.raw = raw
        self.script = ""

    async def run(self, script, timeout=None):
        self.script = script
        return self.raw


class MacEmailFieldParseTests(unittest.TestCase):
    def setUp(self):
        self.original_bridge = server_mac.bridge

    def tearDown(self):
        server_mac.bridge = self.original_bridge

    def test_list_emails_populates_sender_fields(self):
        raw = RECORD_DELIM.join([
            _summary_record("101", "Build broke", "ci@github.com", "Jason Archibald"),
            _summary_record("102", "PR review", "notifications@github.com", "Copilot"),
        ]) + RECORD_DELIM
        server_mac.bridge = FakeBridge(raw)

        result = json.loads(asyncio.run(server_mac.list_emails("inbox", 10)))

        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["sender"], "ci@github.com")
        self.assertEqual(result[0]["sender_name"], "Jason Archibald")
        # Same envelope, different display name — the exact case the fix restores.
        self.assertEqual(result[1]["sender"], "notifications@github.com")
        self.assertEqual(result[1]["sender_name"], "Copilot")

    def test_read_email_by_entry_id_populates_sender_and_recipients(self):
        raw = _full_record(
            "101", "Build broke", "ci@github.com", "Jason Archibald",
            to="me@example.com; ", cc="team@example.com; ",
            body="see logs",
        )
        server_mac.bridge = FakeBridge(raw)

        result = json.loads(asyncio.run(server_mac.read_email(entry_id="101")))

        self.assertEqual(result["sender"], "ci@github.com")
        self.assertEqual(result["sender_name"], "Jason Archibald")
        self.assertEqual(result["to"], "me@example.com;")
        self.assertEqual(result["cc"], "team@example.com;")
        self.assertEqual(result["body"], "see logs")

    def test_read_email_by_subject_search_populates_fields(self):
        raw = _full_record(
            "202", "Feedback wanted", "pm@example.com", "Product Manager",
            to="me@example.com; ", cc="",
            body="thoughts?",
        )
        server_mac.bridge = FakeBridge(raw)

        result = json.loads(
            asyncio.run(server_mac.read_email(subject_search="Feedback"))
        )

        self.assertEqual(result["sender"], "pm@example.com")
        self.assertEqual(result["sender_name"], "Product Manager")
        self.assertEqual(result["to"], "me@example.com;")
        self.assertEqual(result["cc"], "")

    def test_search_emails_populates_sender_fields(self):
        raw = RECORD_DELIM.join([
            _summary_record("301", "Budget report", "cfo@example.com", "Finance"),
        ]) + RECORD_DELIM
        server_mac.bridge = FakeBridge(raw)

        result = json.loads(asyncio.run(server_mac.search_emails("Budget")))

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["sender"], "cfo@example.com")
        self.assertEqual(result[0]["sender_name"], "Finance")


class MacEmailFieldScriptTests(unittest.TestCase):
    """Regression guard on the generated AppleScript access patterns."""

    def setUp(self):
        self.original_bridge = server_mac.bridge

    def tearDown(self):
        server_mac.bridge = self.original_bridge

    def _capture(self, coro_factory):
        fake = FakeBridge("")
        server_mac.bridge = fake
        asyncio.run(coro_factory())
        return fake.script

    def _assert_sender_pattern(self, script):
        # Fixed pattern: bind sender to a variable, then read address/name off it.
        self.assertIn("set snd to sender of m", script)
        self.assertIn("set msender to address of snd", script)
        self.assertIn("set msenderName to name of snd", script)
        # Broken inline pattern must be gone.
        self.assertNotIn("address of sender of m", script)
        self.assertNotIn("name of sender of m", script)

    def _assert_recipient_pattern(self, script):
        # Fixed pattern: resolve nested email address object first.
        self.assertIn("set rea to email address of r", script)
        self.assertIn("address of rea", script)
        # Broken pattern read address straight off the recipient object.
        self.assertNotIn("& address of r &", script)

    def test_list_emails_script_uses_fixed_sender_pattern(self):
        # Use a non-inbox folder so the empty-result UI-scrape fallback (inbox
        # only) doesn't overwrite the captured script.
        script = self._capture(lambda: server_mac.list_emails("sent", 5))
        self._assert_sender_pattern(script)

    def test_read_email_entry_id_script_uses_fixed_patterns(self):
        script = self._capture(lambda: server_mac.read_email(entry_id="101"))
        self._assert_sender_pattern(script)
        self._assert_recipient_pattern(script)

    def test_read_email_subject_search_script_uses_fixed_patterns(self):
        script = self._capture(lambda: server_mac.read_email(subject_search="Feedback"))
        self._assert_sender_pattern(script)
        self._assert_recipient_pattern(script)

    def test_search_emails_script_uses_fixed_sender_pattern(self):
        script = self._capture(lambda: server_mac.search_emails("Budget"))
        self._assert_sender_pattern(script)


if __name__ == "__main__":
    unittest.main()
