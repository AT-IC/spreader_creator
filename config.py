from pathlib import Path


TEMPLATE_PATH = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\DAILY SPREADER -  2026.xlsm"
OUTPUT_DIR = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents"
DOWNLOADS_DIR = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\Downloads"
WMT_URL = "https://wmtscheduler.faa.gov/Views/WorksheetView"
APP_ICON = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\IC-Tools\Spreader Creator\Spreader_Creator_Icon.ico"

CONFIG = {
    "bg_main": "#101C2B", "bg_surface": "#1A2F4C", "bg_surface_alt": "#243B55",
    "bg_input": "#14263B", "bg_log": "#0B1522", "fg_main": "#F0F4F8",
    "fg_muted": "#B8C7D9", "fg_log": "#D8E8F5", "fg_on_accent": "#1A2F4C",
    "accent": "#D99B26", "accent_hover": "#E5AF3C", "accent_pressed": "#B98017",
    "btn_secondary": "#3C619D", "btn_sec_hover": "#5278B5", "border": "#35516F",
    "success": "#63C7AA", "danger": "#FF8FA8", "radius_card": 24,
    "radius_control": 14, "font_family": "Segoe UI",
}


def workspace_path(path: str) -> Path:
    return Path(path).expanduser()
