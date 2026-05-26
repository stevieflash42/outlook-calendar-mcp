import json
import os
import subprocess
import tempfile

import pythoncom
import win32com.client
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageFont

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


SCHEDULE_IMAGE_PATH = os.path.join(tempfile.gettempdir(), "outlook-calendar-schedule.png")

PX_PER_MIN = 1.5
MIN_BLOCK_H = 24
BLOCK_GAP = 2
LEFT_MARGIN = 130
BLOCK_WIDTH = 380
RIGHT_MARGIN = 20
IMG_WIDTH = LEFT_MARGIN + BLOCK_WIDTH + RIGHT_MARGIN

BG = (13, 17, 23)
BUSY_CLR = (218, 54, 51)
POMO_CLR = (35, 134, 54)
BREAK_CLR = (48, 54, 61)
LONG_BREAK_CLR = (83, 68, 23)
TEXT_CLR = (230, 237, 243)
DIM_CLR = (125, 133, 144)


def _load_font(size):
    for name in ["segoeui.ttf", "arial.ttf", "calibri.ttf"]:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _build_day_timeline(date_str, work_start_hour, work_end_hour, slot_duration_minutes, break_duration_minutes, long_break_duration_minutes=15, pomodoros_before_long_break=4):
    items = _get_calendar_items(date_str, date_str)

    busy_events = []
    for item in items:
        if item.BusyStatus == 0:
            continue
        start = _pytime_to_datetime(item.Start)
        end = _pytime_to_datetime(item.End)
        busy_events.append({"start": start, "end": end, "subject": item.Subject})

    busy_events.sort(key=lambda x: x["start"])

    day = datetime.strptime(date_str, "%Y-%m-%d")
    work_start = day.replace(hour=work_start_hour, minute=0)
    work_end = day.replace(hour=work_end_hour, minute=0)
    slot_delta = timedelta(minutes=slot_duration_minutes)
    short_break_delta = timedelta(minutes=break_duration_minutes)
    long_break_delta = timedelta(minutes=long_break_duration_minutes)
    cycle_length = max(1, pomodoros_before_long_break)

    timeline = []
    pomodoro_count = 0
    current = work_start

    def _fill_pomodoros(current, gap_end):
        nonlocal pomodoro_count
        while current + slot_delta <= gap_end:
            pomodoro_count += 1
            pom_end = current + slot_delta
            timeline.append(("slot", current, pom_end, f"Pomodoro {pomodoro_count}"))
            current = pom_end
            is_long = pomodoro_count % cycle_length == 0
            break_delta = long_break_delta if is_long else short_break_delta
            break_type = "long_break" if is_long else "short_break"
            break_label = "Long Break" if is_long else "Short Break"
            if current + break_delta + slot_delta <= gap_end:
                timeline.append((break_type, current, current + break_delta, break_label))
                current = current + break_delta
            elif not is_long and current + short_break_delta + slot_delta <= gap_end:
                timeline.append(("short_break", current, current + short_break_delta, "Short Break"))
                current = current + short_break_delta
            elif current < gap_end:
                break
        return current

    for event in busy_events:
        event_start = max(event["start"], work_start)
        event_end = min(event["end"], work_end)

        if event_end <= current:
            continue
        if event_start >= work_end:
            break

        if event_start > current:
            current = _fill_pomodoros(current, event_start)

        if current < event_start:
            timeline.append(("short_break", current, event_start, "Short Break"))

        timeline.append(("busy", event_start, event_end, event["subject"]))
        current = max(current, event_end)

    if current < work_end:
        _fill_pomodoros(current, work_end)

    return timeline, pomodoro_count, work_start, work_end


def _generate_schedule_image(days_data, slot_duration, break_duration, long_break_duration, pomodoros_before_long_break):
    font_title = _load_font(20)
    font_day = _load_font(16)
    font_label = _load_font(13)
    font_time = _load_font(11)

    DAY_HEADER_H = 40
    DAY_GAP = 20
    top_margin = 50
    bottom_margin = 70

    total_height = top_margin
    for day_info in days_data:
        total_height += DAY_HEADER_H
        for _, start, end, _ in day_info["timeline"]:
            duration = int((end - start).total_seconds() / 60)
            total_height += max(MIN_BLOCK_H, duration * PX_PER_MIN) + BLOCK_GAP
        total_height += DAY_GAP

    total_height += bottom_margin
    img = Image.new("RGB", (IMG_WIDTH, int(total_height)), BG)
    draw = ImageDraw.Draw(img)

    total_pomodoros = sum(d["pomodoro_count"] for d in days_data)
    num_days = len(days_data)
    title = days_data[0]["date"] if num_days == 1 else f"{days_data[0]['date']} to {days_data[-1]['date']}"
    draw.text((20, 16), f"Schedule for {title}", fill=TEXT_CLR, font=font_title)

    y = top_margin
    for day_info in days_data:
        date_obj = datetime.strptime(day_info["date"], "%Y-%m-%d")
        day_label = date_obj.strftime("%A, %b %d")
        time_range = f"{day_info['work_start'].strftime('%I:%M %p')} – {day_info['work_end'].strftime('%I:%M %p')}"
        pomo_label = f"  ({day_info['pomodoro_count']} pomodoros)"

        draw.text((20, y + 10), day_label, fill=TEXT_CLR, font=font_day)
        draw.text((20 + draw.textlength(day_label, font=font_day) + 10, y + 14), time_range, fill=DIM_CLR, font=font_time)
        draw.text((20 + draw.textlength(day_label, font=font_day) + 10 + draw.textlength(time_range, font=font_time) + 8, y + 14), pomo_label, fill=POMO_CLR, font=font_time)
        y += DAY_HEADER_H

        entries = day_info["timeline"]
        for i, (entry_type, start, end, label) in enumerate(entries):
            duration = int((end - start).total_seconds() / 60)
            h = max(MIN_BLOCK_H, duration * PX_PER_MIN)
            color = BUSY_CLR if entry_type == "busy" else POMO_CLR if entry_type == "slot" else LONG_BREAK_CLR if entry_type == "long_break" else BREAK_CLR if entry_type == "short_break" else BREAK_CLR

            draw.rounded_rectangle(
                [LEFT_MARGIN, y, LEFT_MARGIN + BLOCK_WIDTH, y + h],
                radius=5, fill=color,
            )

            start_str = start.strftime("%I:%M %p")
            end_str = end.strftime("%I:%M %p")
            next_starts_at_end = i + 1 < len(entries) and entries[i + 1][1] == end
            draw.text((15, y + 4), start_str, fill=DIM_CLR, font=font_time)
            if h >= 50 and not next_starts_at_end:
                draw.text((15, y + h - 16), end_str, fill=DIM_CLR, font=font_time)

            if h >= MIN_BLOCK_H and label:
                max_chars = BLOCK_WIDTH // 8
                display = (label[:max_chars - 3] + "...") if len(label) > max_chars else label
                draw.text((LEFT_MARGIN + 10, y + (h - 13) / 2), display, fill=TEXT_CLR, font=font_label)

            y += h + BLOCK_GAP

        y += DAY_GAP

    legend_y = y
    legend_x = 20
    for legend_label, legend_color in [("Busy", BUSY_CLR), ("Pomodoro", POMO_CLR), ("Short Break", BREAK_CLR), ("Long Break", LONG_BREAK_CLR)]:
        draw.rounded_rectangle(
            [legend_x, legend_y, legend_x + 12, legend_y + 12],
            radius=2, fill=legend_color,
        )
        draw.text((legend_x + 18, legend_y - 1), legend_label, fill=DIM_CLR, font=font_label)
        legend_x += 100

    summary = f"{total_pomodoros} pomodoro(s) total ({slot_duration} min work, {break_duration} min break, {long_break_duration} min long break every {pomodoros_before_long_break})"
    draw.text((20, legend_y + 22), summary, fill=DIM_CLR, font=font_label)

    img.save(SCHEDULE_IMAGE_PATH)
    return SCHEDULE_IMAGE_PATH


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
    start_date: str,
    end_date: str = "",
    slot_duration_minutes: int = 25,
    break_duration_minutes: int = 5,
    long_break_duration_minutes: int = 15,
    pomodoros_before_long_break: int = 4,
    work_start_hour: int = 9,
    work_end_hour: int = 17,
) -> str:
    """Find free time slots on one or more days with a visual timeline.

    Args:
        start_date: Start date in YYYY-MM-DD format
        end_date: End date in YYYY-MM-DD format (optional, defaults to start_date)
        slot_duration_minutes: Work slot duration in minutes (default: 25)
        break_duration_minutes: Short break between slots in minutes (default: 5)
        long_break_duration_minutes: Long break after every Nth pomodoro in minutes (default: 15)
        pomodoros_before_long_break: Number of pomodoros before a long break (default: 4)
        work_start_hour: Work day start hour in 24h format (default: 9)
        work_end_hour: Work day end hour in 24h format (default: 17)
    """
    if not end_date:
        end_date = start_date

    current_day = datetime.strptime(start_date, "%Y-%m-%d")
    last_day = datetime.strptime(end_date, "%Y-%m-%d")

    days_data = []
    lines = []
    total_pomodoros = 0

    while current_day <= last_day:
        date_str = current_day.strftime("%Y-%m-%d")
        timeline, pomodoro_count, work_start, work_end = _build_day_timeline(
            date_str, work_start_hour, work_end_hour,
            slot_duration_minutes, break_duration_minutes,
            long_break_duration_minutes, pomodoros_before_long_break,
        )

        days_data.append({
            "date": date_str,
            "timeline": timeline,
            "pomodoro_count": pomodoro_count,
            "work_start": work_start,
            "work_end": work_end,
        })

        total_pomodoros += pomodoro_count
        fmt = "%I:%M %p"
        day_label = current_day.strftime("%A, %b %d")
        lines.append(f"{day_label} ({work_start.strftime(fmt)} - {work_end.strftime(fmt)}):")
        lines.append("")

        for entry_type, start, end, label in timeline:
            s = start.strftime(fmt)
            e = end.strftime(fmt)
            if entry_type == "busy":
                lines.append(f"  {s} - {e}  ██ {label}")
            elif entry_type == "slot":
                lines.append(f"  {s} - {e}  ░░ {label}")
            elif entry_type == "long_break":
                lines.append(f"  {s} - {e}  ·· long break")
            elif entry_type == "short_break":
                lines.append(f"  {s} - {e}  ·· short break")

        lines.append(f"  ({pomodoro_count} pomodoros)")
        lines.append("")

        current_day += timedelta(days=1)

    lines.append(f"Total: {total_pomodoros} pomodoro(s) ({slot_duration_minutes} min work, {break_duration_minutes} min break, {long_break_duration_minutes} min long break every {pomodoros_before_long_break})")

    _generate_schedule_image(days_data, slot_duration_minutes, break_duration_minutes, long_break_duration_minutes, pomodoros_before_long_break)
    subprocess.Popen(["code", SCHEDULE_IMAGE_PATH], shell=True, creationflags=subprocess.DETACHED_PROCESS)

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