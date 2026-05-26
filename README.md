# Outlook Calendar MCP Server

An MCP (Model Context Protocol) server that gives Claude Code read/write access to your Outlook calendar via COM automation. Built to find free time slots and book events without requiring IT/admin involvement.

## Prerequisites

- Windows 10/11
- Python 3.12+
- **Classic Outlook** (not "New Outlook") installed with your account configured
- [Claude Code](https://claude.ai/download)

> **Why classic Outlook?** New Outlook is a web app and doesn't expose calendars through COM/MAPI. Classic Outlook must be installed and have your accounts added, but you can continue using New Outlook as your daily client — just not at the same time (see Troubleshooting).

## Setup

1. Clone this repo:
   ```
   git clone <repo-url> C:\Projects\outlook-calendar-mcp
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Open classic Outlook and add your email account(s). You can close it afterward — COM will launch it automatically when needed.

4. Register the MCP server with Claude Code:
   ```
   claude mcp add -s user outlook-calendar -- python "C:\Projects\outlook-calendar-mcp\server.py"
   ```

5. Restart Claude Code. The calendar tools will be available.

## Configuration

On first use, run `list_calendars` to discover available calendar folders, then `set_active_calendars` to choose which ones to include when querying events. Configuration is saved in `config.json`.

## Tools

| Tool | Description |
|---|---|
| `list_calendars` | List all calendar folders across all Outlook accounts |
| `set_active_calendars` | Choose which calendars to query |
| `list_events` | List events in a date range |
| `find_free_slots` | Find available time slots on a given date (default: 25-minute slots, 9 AM - 5 PM) |
| `create_event` | Create a calendar event |
| `delete_event` | Delete an event by ID |

## Notes

- Classic Outlook does not need to be running — COM will launch it automatically.
- Events marked as "Free" are ignored when calculating free slots.
- Recurring events are properly expanded to their actual occurrence dates.
- `config.json` is machine-specific (different accounts/calendar paths per machine) and is gitignored.

## Troubleshooting

**`Server execution failed` / `0x80080005` COM error:**

This means COM cannot activate Outlook. Common causes:

- **New Outlook and Classic Outlook open simultaneously.** Close New Outlook, kill Classic Outlook (Task Manager), then retry. COM will relaunch Classic Outlook in a clean state.
- **Elevation mismatch.** If Outlook was launched as administrator but your editor/Claude Code is running as a normal user (or vice versa), COM will fail. Both processes must run at the same integrity level — close Outlook, ensure nothing is running as admin, and retry.

**Schedule image doesn't open in VS Code:**

If using `claude mcp add` manually, do not set `"env": {}` in the MCP server config — this strips the system PATH and prevents the `code` CLI from being found. Omit the `env` key entirely so the server inherits the parent process's environment.
