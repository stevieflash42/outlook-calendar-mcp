import json
import os

import pythoncom
import win32com.client
from datetime import datetime, timedelta

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("outlook-calendar")

BUSY_STATUS_LABELS = ["Free", "Tentative", "Busy", "Out of Office", "Working Elsewhere"]
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def _get_outlook():
    pythoncom.CoInitialize()
    return win32com.client.Dispatch("Outlook.Application")


def _pytime_to_datetime(pytime) -> datetime:
    return datetime(
        pytime.year, pytime.month, pytime.day,
        pytime.hour, pytime.minute, pytime.second,
    )


def _load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    return {"active_calendars": []}


def _save_config(config: dict) -> None:
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


def _find_calendar_folders(folder, path=""):
    calendars = []
    try:
        if folder.DefaultItemType == 1:
            current_path = f"{path}{folder.Name}" if path else folder.Name
            calendars.append({
                "name": folder.Name,
                "path": current_path,
                "entry_id": folder.EntryID,
                "store_id": folder.StoreID,
            })
    except Exception:
        pass
    try:
        for i in range(folder.Folders.Count):
            subfolder = folder.Folders.Item(i + 1)
            sub_path = f"{path}{folder.Name}/" if path else f"{folder.Name}/"
            calendars.extend(_find_calendar_folders(subfolder, sub_path))
    except Exception:
        pass
    return calendars


def _get_all_calendars():
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    calendars = []
    for store in namespace.Stores:
        root = store.GetRootFolder()
        calendars.extend(_find_calendar_folders(root))
    return calendars


def _get_active_calendar_folders():
    config = _load_config()
    active_paths = config.get("active_calendars", [])

    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")

    if not active_paths:
        return [namespace.GetDefaultFolder(9)]

    all_calendars = _get_all_calendars()
    folders = []
    for cal in all_calendars:
        if cal["path"] in active_paths:
            folder = namespace.GetFolderFromID(cal["entry_id"], cal["store_id"])
            folders.append(folder)

    if not folders:
        return [namespace.GetDefaultFolder(9)]

    return folders


def _get_calendar_items(start_date: str, end_date: str):
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")

    restriction = (
        f"[Start] >= '{start_dt.month}/{start_dt.day}/{start_dt.year} 12:00 AM'"
        f" AND [Start] <= '{end_dt.month}/{end_dt.day}/{end_dt.year} 11:59 PM'"
    )

    all_items = []
    for folder in _get_active_calendar_folders():
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        for item in items.Restrict(restriction):
            all_items.append(item)

    all_items.sort(key=lambda item: _pytime_to_datetime(item.Start))
    return all_items


@mcp.tool()
def list_events(start_date: str, end_date: str) -> str:
    """List calendar events in a date range.

    Args:
        start_date: Start date in YYYY-MM-DD format
        end_date: End date in YYYY-MM-DD format
    """
    restricted = _get_calendar_items(start_date, end_date)

    results = []
    for item in restricted:
        start = _pytime_to_datetime(item.Start)
        end = _pytime_to_datetime(item.End)
        status = BUSY_STATUS_LABELS[item.BusyStatus] if item.BusyStatus < len(BUSY_STATUS_LABELS) else "Unknown"
        results.append(
            f"- {item.Subject}\n"
            f"  {start.strftime('%Y-%m-%d %I:%M %p')} to {end.strftime('%I:%M %p')}\n"
            f"  Location: {item.Location or '(none)'}\n"
            f"  Status: {status}\n"
            f"  ID: {item.EntryID}"
        )

    if not results:
        return "No events found in the specified range."

    return f"Found {len(results)} event(s):\n\n" + "\n\n".join(results)


@mcp.tool()
def find_free_slots(
    date: str,
    slot_duration_minutes: int = 25,
    work_start_hour: int = 9,
    work_end_hour: int = 17,
) -> str:
    """Find free time slots on a given date.

    Args:
        date: Date in YYYY-MM-DD format
        slot_duration_minutes: Minimum slot duration in minutes (default: 120)
        work_start_hour: Work day start hour in 24h format (default: 9)
        work_end_hour: Work day end hour in 24h format (default: 17)
    """
    restricted = _get_calendar_items(date, date)

    busy_periods = []
    for item in restricted:
        if item.BusyStatus == 0:
            continue
        start = _pytime_to_datetime(item.Start)
        end = _pytime_to_datetime(item.End)
        busy_periods.append((start, end))

    busy_periods.sort(key=lambda x: x[0])

    day = datetime.strptime(date, "%Y-%m-%d")
    work_start = day.replace(hour=work_start_hour, minute=0)
    work_end = day.replace(hour=work_end_hour, minute=0)
    slot_delta = timedelta(minutes=slot_duration_minutes)

    free_slots = []
    current = work_start

    for busy_start, busy_end in busy_periods:
        if busy_end <= current:
            continue
        if busy_start >= work_end:
            break
        if busy_start > current:
            gap_end = min(busy_start, work_end)
            while current + slot_delta <= gap_end:
                free_slots.append((current, current + slot_delta))
                current = current + slot_delta
        current = max(current, busy_end)

    while current + slot_delta <= work_end:
        free_slots.append((current, current + slot_delta))
        current = current + slot_delta

    if not free_slots:
        return f"No free {slot_duration_minutes}-minute slots on {date}."

    lines = [f"Free {slot_duration_minutes}-minute slots on {date}:"]
    for i, (s, e) in enumerate(free_slots, 1):
        lines.append(f"  {i}. {s.strftime('%I:%M %p')} - {e.strftime('%I:%M %p')}")
    return "\n".join(lines)


@mcp.tool()
def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
) -> str:
    """Create a calendar event.

    Args:
        subject: Event title
        start: Start time in 'YYYY-MM-DD HH:MM' format (24h)
        end: End time in 'YYYY-MM-DD HH:MM' format (24h)
        location: Optional location
        body: Optional description
    """
    outlook = _get_outlook()
    appointment = outlook.CreateItem(1)
    appointment.Subject = subject
    appointment.Start = start
    appointment.End = end
    if location:
        appointment.Location = location
    if body:
        appointment.Body = body
    appointment.Save()
    return f"Created '{subject}' from {start} to {end}."


@mcp.tool()
def delete_event(entry_id: str) -> str:
    """Delete a calendar event by its EntryID (from list_events output).

    Args:
        entry_id: The EntryID of the event to delete
    """
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    item = namespace.GetItemFromID(entry_id)
    subject = item.Subject
    item.Delete()
    return f"Deleted '{subject}'."


@mcp.tool()
def list_calendars() -> str:
    """List all available Outlook calendar folders across all accounts."""
    calendars = _get_all_calendars()
    config = _load_config()
    active_paths = config.get("active_calendars", [])

    if not calendars:
        return "No calendar folders found."

    lines = ["Available calendars:"]
    for cal in calendars:
        marker = " [ACTIVE]" if cal["path"] in active_paths else ""
        lines.append(f"  - {cal['path']}{marker}")
    return "\n".join(lines)


@mcp.tool()
def set_active_calendars(calendar_paths: list[str]) -> str:
    """Set which calendars to include when querying events.

    Args:
        calendar_paths: List of calendar paths from list_calendars output
    """
    all_calendars = _get_all_calendars()
    valid_paths = {cal["path"] for cal in all_calendars}

    invalid = [p for p in calendar_paths if p not in valid_paths]
    if invalid:
        return f"Invalid calendar paths: {', '.join(invalid)}\nRun list_calendars to see available paths."

    config = _load_config()
    config["active_calendars"] = calendar_paths
    _save_config(config)

    return "Active calendars set to:\n" + "\n".join(f"  - {p}" for p in calendar_paths)


if __name__ == "__main__":
    mcp.run()