# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook and start a Claude Code session.** That's it — tools are available immediately.

## How It Works — Platform Routing

When the server starts, it checks which operating system it is running on and takes one of two paths:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (29 tools)    │    │  Bridge          │
              │                │    │  (22 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks — but does not expose rules, categories, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations with identical tool names and signatures, so MCP clients see the same interface regardless of platform. Tools that are not available on a given platform are simply not registered.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `send_email` | yes | yes | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | yes | yes | List recent emails from any folder, with optional unread filter |
| `read_email` | yes | yes | Read full email content by entry ID or subject search |
| `search_emails` | yes | yes | Full-text search across email subjects and bodies |
| `reply_email` | yes | yes | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | yes | yes | Mark a specific email as read |
| `mark_as_unread` | yes | yes | Mark a specific email as unread |
| `move_email` | yes | yes | Move an email to Archive, Trash, or any folder |
| `list_folders` | yes | yes | Browse the folder hierarchy with item counts |

### Calendar

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_events` | yes | yes | List upcoming events within a date range |
| `get_event` | yes | yes | Read full event details by entry ID |
| `create_event` | yes | yes | Create a personal calendar appointment |
| `create_meeting` | yes | yes | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | Modify an existing event's subject, time, location, etc. |
| `delete_event` | yes | yes | Delete an appointment or cancel a meeting |
| `respond_to_meeting` | yes | — | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | yes | yes | Search calendar events by keyword within a date range |

### Tasks

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_tasks` | yes | yes | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | Read full task details including body and completion status |
| `create_task` | yes | yes | Create a new task with subject, due date, importance |
| `complete_task` | yes | yes | Mark a task as complete |
| `delete_task` | yes | yes | Remove a task |

### Attachments

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_attachments` | yes | yes | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | Download an attachment to a local directory |

### Categories, Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (MAPI property accessors, the Rules object model, and the Categories collection) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | — | List all available color categories in Outlook |
| `set_category` | yes | — | Set or clear categories on any email, event, or task |
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

**Total: 29 tools on Windows, 22 tools on macOS.**

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters.
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.

## Install from Source

### Windows

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

Windows-only examples:

- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection → routes to correct server
    server.py                # Windows MCP server (29 tools, COM automation)
    server_mac.py            # macOS MCP server (22 tools, AppleScript)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      errors.py              # COM error formatting (Windows)
      applescript_helpers.py # AppleScript escaping, date formatting (macOS)
  tests/
    phase1_com_test.py       # Email COM validation
    phase3_mcp_test.py       # Email MCP test
    calendar_com_test.py     # Calendar COM validation
    calendar_mcp_test.py     # Calendar MCP test
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
