from pathlib import Path
from typing import Callable, List
import re

from bs4 import BeautifulSoup
from models import ParsedData


def _log(log_cb: Callable[[str], None], message: str) -> None:
    log_cb(message)


def parse_html_data(filepath: str, log_cb: Callable[[str], None] = lambda _: None) -> ParsedData:
    path = Path(filepath.strip('"').strip("'"))
    _log(log_cb, f"Opening file: {path}")
    with path.open("r", encoding="utf-8", errors="ignore") as handle:
        soup = BeautifulSoup(handle.read(), "html.parser")

    ot_dict = {}
    _log(log_cb, "--- SCANNING OVERTIME TABLE ---")
    for row in soup.find_all("tr"):
        cols = row.find_all("td")
        if len(cols) != 5 or cols[1].get_text(strip=True).casefold() not in {"scheduled", "call in", "holdover"}:
            continue
        init = cols[0].get_text(strip=True).upper()
        hours_text = cols[2].get_text(strip=True)
        reason = cols[3].get_text(strip=True).upper()
        if hours_text.isdigit():
            entry = ot_dict.setdefault(init, {"hrs": 0, "fos": 0})
            entry["hrs"] += int(hours_text)
            match = re.search(r"(\d+)\s*FOS", reason)
            entry["fos"] += int(match.group(1)) if match else 0
            _log(log_cb, f"OT Captured -> {init}: {hours_text} hr(s) | {entry['fos']} FOS total")

    leave_list = []
    _log(log_cb, "--- SCANNING LEAVE/REQUESTS ---")
    for table_id in ("tblShiftChange", "tblXTECTE", "tblSupRequests"):
        table = soup.find("table", id=table_id)
        if not table:
            continue
        for row in table.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) < 6:
                continue
            init, leave_type, leave_from, leave_to = [cols[i].get_text(strip=True) for i in range(4)]
            status = cols[5].get_text(strip=True)
            normalized_status = status.casefold()
            normalized_type = leave_type.casefold()
            if "approved" in normalized_status or "shift chg" in normalized_type or "shift swap" in normalized_type:
                continue
            addon = ""
            if leave_from and leave_from.casefold() != "shift":
                addon = f" ({leave_from} - {leave_to})" if leave_to else f" ({leave_from})"
            leave_list.append(f"{init} - {leave_type}{addon}")

    notes_list = []
    _log(log_cb, "--- SCANNING SHIFT NOTES ---")
    notes_table = soup.find("table", id="ShiftNotesForDay")
    if notes_table:
        for row in notes_table.find_all("tr"):
            cols = row.find_all("td")
            link = cols[1].find("a", title=True) if len(cols) >= 2 else None
            if not link:
                continue
            note = link.get_text(strip=True)
            if not note or note == "<Add Shift Notes>":
                continue
            cleaned = re.sub(r"^[A-Za-z]{2,3}\s*-\s*", "", note)
            if re.match(r"^\d{2}\s*-", cleaned):
                continue
            notes_list.append(cleaned)

    shifts = []
    _log(log_cb, "--- PROCESSING SHIFTS ---")
    ignored = {"tblShiftChange", "tblXTECTE", "tblSupRequests", "ShiftNotesForDay"}
    for row in soup.find_all("tr"):
        cols = row.find_all("td")
        parent = row.find_parent("table")
        if len(cols) < 4 or (parent and parent.get("id") in ignored):
            continue
        time_match = re.match(r"^(\d{2,4})([A-Z]*)$", cols[0].get_text(strip=True).replace("\xa0", ""))
        if not time_match:
            continue
        raw_time, orig_code = time_match.groups()
        start = raw_time.zfill(4) if len(raw_time) == 3 else (raw_time + "00" if len(raw_time) <= 2 else raw_time)
        for idx, col in enumerate(cols[2:4]):
            is_dev = idx == 1
            for word in col.get_text(separator=" ", strip=True).replace("\xa0", " ").split():
                match = re.match(r"^([A-Z]{2})(\$?)$", word)
                if not match:
                    continue
                initials, has_dollar = match.group(1), bool(match.group(2))
                if initials in {"OF", "OT"}:
                    continue
                ot_hours = ot_dict.get(initials, {}).get("hrs", 0)
                fos_hours = ot_dict.get(initials, {}).get("fos", 0)
                if is_dev:
                    final_code = "T"
                else:
                    final_code = orig_code
                    if has_dollar:
                        final_code = "$$" if orig_code == "A" or ot_hours > 8 else "$"
                    elif not orig_code and 0 < ot_hours < 8:
                        final_code = "$$"
                shifts.append({"initials": initials, "start": start, "orig_code": orig_code,
                               "code": final_code, "ot": ot_hours, "fos": fos_hours})
    return ParsedData(shifts, leave_list, notes_list)


def preview_text(data: ParsedData) -> str:
    lines = [f"Shifts: {len(data.shifts)}", f"Flagged requests: {len(data.leave_list)}",
             f"Shift notes: {len(data.notes_list)}", "", "Shift preview:"]
    lines.extend(f"  {s['initials']} | {s['start']} | {s['code'] or 'Standard'}" for s in data.shifts[:20])
    if len(data.shifts) > 20:
        lines.append(f"  ... and {len(data.shifts) - 20} more")
    return "\n".join(lines)
