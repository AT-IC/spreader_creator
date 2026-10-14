from typing import Tuple


def to_num(value: str):
    if not value:
        return ""
    return int(value) if value.isdigit() else value


def subtract_hours(time_str: str, hours_to_sub: int) -> str:
    if time_str in {"MID", "2400"}:
        time_str = "0000"
    try:
        total_minutes = int(time_str[:2]) * 60 + int(time_str[2:4])
        total_minutes = (total_minutes - hours_to_sub * 60) % (24 * 60)
        return f"{total_minutes // 60:02d}{total_minutes % 60:02d}"
    except (TypeError, ValueError):
        return time_str


def add_hours(hour: int, minute: int, hours_to_add: int) -> str:
    total_minutes = (hour * 60 + minute + hours_to_add * 60) % (24 * 60)
    if total_minutes == 0 and hours_to_add > 0:
        return "2400"
    return f"{total_minutes // 60:02d}{total_minutes % 60:02d}"


def calculate_times(start_time: str, final_code: str, orig_code: str,
                    ot_hours: int = 0, fos_hours: int = 0) -> Tuple[str, str]:
    if start_time in {"2230", "2300", "2330", "2400"}:
        return ("", "MID") if final_code == "T" else ("MID", "")
    adjusted_start = subtract_hours(start_time, fos_hours)
    try:
        start_h, start_m = int(adjusted_start[:2]), int(adjusted_start[2:4])
    except (TypeError, ValueError):
        return "", ""
    if final_code == "T":
        shift_len = ot_hours if ot_hours >= 8 else 8 + fos_hours
        return "", add_hours(start_h, start_m, shift_len)
    if final_code == "S":
        return add_hours(start_h, start_m, 7), ""
    if final_code == "N":
        return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 9)
    if final_code == "A":
        return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 10)
    if final_code == "$$":
        if orig_code == "A" or ot_hours == 0:
            return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 10)
        if ot_hours > 8:
            return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, ot_hours)
        return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 8 + ot_hours)
    if final_code == "$":
        return add_hours(start_h, start_m, 8), ""
    return add_hours(start_h, start_m, 8), ""


def sort_chronological(shift: dict) -> int:
    start = int(shift["start"])
    return start - 2400 if start >= 2200 else start
