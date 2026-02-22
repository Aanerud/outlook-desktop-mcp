# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

**Turn your running Outlook Desktop into an MCP server with 29 tools.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+ on Windows):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook Desktop (Classic) and start a Claude Code session.** That's it — 29 tools are available immediately.

## How It Works

```
Claude Code / Claude Desktop / Any MCP Client
    |
    | stdio (JSON-RPC)
    v
outlook-desktop-mcp (Python)
    |
    | COM automation via Outlook Object Model (MSOUTL.OLB)
    v
Outlook Desktop (Classic) — OUTLOOK.EXE
    |
    | Your existing authenticated session
    v
Exchange Online / Microsoft 365 / On-Premises Exchange
```

The server uses Windows COM automation to talk directly to the running `OUTLOOK.EXE` process. It inherits whatever authentication Outlook already has — your M365 account, on-prem Exchange, or even personal Outlook.com accounts. No additional credentials or API keys are needed.

Internally, the server runs a dedicated COM thread (Single-Threaded Apartment) that holds the `Outlook.Application` object. The async MCP event loop dispatches tool calls to this thread via a queue, keeping COM threading rules respected and the MCP protocol non-blocking.

## Requirements

- **Windows** — COM automation is Windows-only
- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

## Available Tools (29)

All tool descriptions are optimized for LLM tool discovery — Claude understands exactly how to use each one, what arguments to pass, and what to expect back.

### Email (9 tools)

| Tool | Description |
|------|-------------|
| `send_email` | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | List recent emails from any folder, with optional unread filter |
| `read_email` | Read full email content by entry ID or subject search |
| `search_emails` | Full-text search across email subjects and bodies |
| `reply_email` | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | Mark a specific email as read |
| `mark_as_unread` | Mark a specific email as unread |
| `move_email` | Move an email to Archive, Trash, or any folder |
| `list_folders` | Browse the complete folder hierarchy with item counts |

### Calendar (8 tools)

| Tool | Description |
|------|-------------|
| `list_events` | List upcoming events with recurring occurrence support |
| `get_event` | Read full event details by entry ID |
| `create_event` | Create a personal calendar appointment |
| `create_meeting` | Create a meeting and send invitations to attendees |
| `update_event` | Modify an existing event's subject, time, location, etc. |
| `delete_event` | Delete an appointment or cancel a meeting (sends notices) |
| `respond_to_meeting` | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | Search calendar events by keyword within a date range |

### Tasks (5 tools)

| Tool | Description |
|------|-------------|
| `list_tasks` | List pending or completed tasks, sorted by due date |
| `get_task` | Read full task details including body and completion status |
| `create_task` | Create a new task with subject, due date, importance |
| `complete_task` | Mark a task as complete (100%) |
| `delete_task` | Remove a task |

### Attachments (2 tools)

| Tool | Description |
|------|-------------|
| `list_attachments` | List all attachments on an email or calendar event |
| `save_attachment` | Download an attachment to a local directory |

### Categories (2 tools)

| Tool | Description |
|------|-------------|
| `list_categories` | List all available color categories in Outlook |
| `set_category` | Set or clear categories on any email, event, or task |

### Rules (2 tools)

| Tool | Description |
|------|-------------|
| `list_rules` | List all mail rules with enabled/disabled status |
| `toggle_rule` | Enable or disable a mail rule by name |

### Out of Office (1 tool)

| Tool | Description |
|------|-------------|
| `get_out_of_office` | Check whether Out of Office auto-reply is on or off |

## Install from Source

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
| Works offline / cached | No | Yes (reads from OST cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    server.py              # MCP server + all 29 tool definitions
    com_bridge.py          # Async-to-COM threading bridge
    tools/
      _folder_constants.py # Outlook enums and constants
    utils/
      formatting.py        # Email, event, and task data extraction
      errors.py            # COM error formatting
  tests/
    phase1_com_test.py     # Email COM validation
    phase3_mcp_test.py     # Email MCP test
    calendar_com_test.py   # Calendar COM validation
    calendar_mcp_test.py   # Calendar MCP test
    extras_com_test.py     # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py     # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd  # Windows launcher script
  pyproject.toml
```

## License

See [LICENSE](LICENSE) file.
