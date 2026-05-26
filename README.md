# Outlook Calendar MCP Server

An MCP (Model Context Protocol) server that gives Claude Code read/write access to your Outlook calendar via COM automation. Built to find free time slots and book events without requiring IT/admin involvement.

## Prerequisites

- Windows 10/11
- Python 3.12+
- **Classic Outlook** (not "New Outlook") installed and running with your account configured
- [Claude Code](https://claude.ai/download)

> **Why classic Outlook?** New Outlook is a web app and doesn't expose calendars through COM/MAPI. Classic Outlook must be installed and have your accounts added, but you can continue using New Outlook as your daily client.

## Setup

1. Clone this repo:
   ```
   git clone <repo-url> C:\Projects\outlook-calendar-mcp
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Open classic Outlook and add your email account(s).

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

- Events marked as "Free" are ignored when calculating free slots.
- Recurring events are properly expanded to their actual occurrence dates.
- `config.json` is machine-specific (different accounts/calendar paths per machine) and is gitignored.
