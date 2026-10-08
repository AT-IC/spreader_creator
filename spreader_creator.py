import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
from datetime import datetime, timedelta
import openpyxl
from bs4 import BeautifulSoup
import traceback
import os
import webbrowser
import calendar

# ==========================================
# 1. CONFIGURATION & THEMES
# ==========================================
TEMPLATE_PATH = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\DAILY SPREADER -  2026.xlsm"
OUTPUT_DIR = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\\"
DOWNLOADS_DIR = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\Downloads"
WMT_URL = "https://wmtscheduler.faa.gov/Views/WorksheetView"
APP_ICON = r"\\ITPAATPNFP001.ad.faa.gov\TPA-Data\Air Traffic\Watch Desk Documents\IC-Tools\Spreader Creator\Spreader_Creator_Icon.ico"

COLORS = {
    "bg_main": "#202020",
    "fg_main": "#ffffff",
    "bg_input": "#333333",
    "bg_log": "#1e1e1e",
    "fg_log": "#00d084",
    "btn_primary": "#005A9C",
    "btn_hover": "#0073C4",
    "btn_secondary": "#444444",
    "btn_sec_hover": "#555555",
    "btn_help": "#117711",
    "btn_help_hover": "#119911"
}

# ==========================================
# 2. CORE BUSINESS LOGIC & TIME MATH
# ==========================================
def to_num(val):
    if not val: return ""
    if val.isdigit(): return int(val)
    return val

def subtract_hours(time_str, hours_to_sub):
    if time_str in ["MID", "2400"]: time_str = "0000"
    try:
        h = int(time_str[:2])
        m = int(time_str[2:4])
        new_h = h - hours_to_sub
        if new_h < 0: new_h += 24
        return f"{new_h:02d}{m:02d}"
    except ValueError:
        return time_str

def add_hours(h, m, hours_to_add):
    new_h = h + hours_to_add
    while new_h >= 24:
        if new_h == 24 and m == 0: return "2400"
        new_h -= 24
    return f"{new_h:02d}{m:02d}"

def calculate_times(start_time, final_code, orig_code, ot_hours=0, fos_hours=0):
    if start_time in ["2230", "2300", "2330", "2400"]:
        return ("", "MID") if final_code == "T" else ("MID", "")
    
    # FOS Start time adjustment for calculation purposes only
    adjusted_start = subtract_hours(start_time, fos_hours)
    
    try:
        start_h, start_m = int(adjusted_start[:2]), int(adjusted_start[2:4])
    except ValueError:
        return "", ""

    if final_code == "T":       # Dev shift
        shift_len = ot_hours if ot_hours >= 8 else 8 + fos_hours
        return "", add_hours(start_h, start_m, shift_len)
    elif final_code == "S":     # Standard
        return add_hours(start_h, start_m, 7), ""
    elif final_code == "N":     # 9-hour
        return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 9)
    elif final_code == "A":     # 10-hour
        return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 10)
    elif final_code == "$$":    # Extended OT
        if orig_code == "A" or ot_hours == 0:
            return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 10)
        elif ot_hours > 8:
            return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, ot_hours)
        else:
            return add_hours(start_h, start_m, 8), add_hours(start_h, start_m, 8 + ot_hours)
    elif final_code == "$":     # Standard OT
        return add_hours(start_h, start_m, 8), ""
    else:
        return add_hours(start_h, start_m, 8), ""

def sort_chronological(shift):
    t = int(shift['start'])
    return t - 2400 if t >= 2200 else t


# ==========================================
# 3. DATA EXTRACTION (HTML PARSING)
# ==========================================
def parse_html_data(filepath, log_cb):
    filepath = filepath.strip('"').strip("'")
    log_cb(f"Opening file: {filepath}")
    
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    
    # --- 1. OVERTIME ---
    ot_dict = {}
    log_cb("--- SCANNING OVERTIME TABLE ---")
    for row in soup.find_all('tr'):
        cols = row.find_all('td')
        if len(cols) == 5 and cols[1].get_text(strip=True) in ["Scheduled", "Call In", "Holdover"]:
            init = cols[0].get_text(strip=True)
            hrs_text = cols[2].get_text(strip=True)
            reason_text = cols[3].get_text(strip=True).upper()
            
            if hrs_text.isdigit():
                hrs = int(hrs_text)
                fos_match = re.search(r'(\d+)\s*FOS', reason_text)
                fos = int(fos_match.group(1)) if fos_match else 0
                
                if init not in ot_dict: ot_dict[init] = {'hrs': 0, 'fos': 0}
                ot_dict[init]['hrs'] += hrs
                ot_dict[init]['fos'] += fos
                log_cb(f"OT Captured -> {init}: {hrs} hr(s) | {fos} FOS")

    # --- 2. REQUESTS ---
    leave_list = []
    log_cb("--- SCANNING LEAVE/REQUESTS ---")
    for table_id in ['tblShiftChange', 'tblXTECTE', 'tblSupRequests']:
        req_table = soup.find('table', id=table_id)
        if not req_table: continue
            
        for row in req_table.find_all('tr'):
            cols = row.find_all('td')
            if len(cols) >= 6:
                init, lv_type, lv_from, lv_to = [cols[i].get_text(strip=True) for i in range(4)]
                status = cols[5].get_text(strip=True)
                
                if "Approved" not in status and "Shift Chg" not in lv_type and "Shift Swap" not in lv_type:
                    time_addon = ""
                    if lv_from and lv_from.lower() != "shift":
                        time_addon = f" ({lv_from} - {lv_to})" if lv_to else f" ({lv_from})"
                    leave_entry = f"{init} - {lv_type}{time_addon}"
                    leave_list.append(leave_entry)
                    log_cb(f"Flagged -> {leave_entry} (Status: {status})")

    # --- 3. NOTES ---
    notes_list = []
    log_cb("--- SCANNING SHIFT NOTES ---")
    notes_table = soup.find('table', id='ShiftNotesForDay')
    if notes_table:
        for row in notes_table.find_all('tr'):
            cols = row.find_all('td')
            if len(cols) >= 2 and cols[1].find('a', title=True):
                note_text = cols[1].find('a', title=True).get_text(strip=True)
                if note_text and note_text != "<Add Shift Notes>":
                    cleaned_note = re.sub(r'^[A-Za-z]{2,3}\s*-\s*', '', note_text)
                    
                    # Filter out WMT auto-generated L&A/CruArt notes (e.g., "65 - OT..." or "51 - LOT...")
                    if re.match(r'^\d{2}\s*-', cleaned_note):
                        log_cb(f"Ignored System Note -> {cleaned_note}")
                    else:
                        notes_list.append(cleaned_note)
                        log_cb(f"Note -> {cleaned_note}")

    # --- 4. SHIFTS ---
    shifts = []
    log_cb("--- PROCESSING SHIFTS ---")
    ignore_tables = ['tblShiftChange', 'tblXTECTE', 'tblSupRequests', 'ShiftNotesForDay']
    
    for row in soup.find_all('tr'):
        cols = row.find_all('td')
        if len(cols) < 4: continue
        
        parent = row.find_parent('table')
        if parent and parent.get('id') in ignore_tables: continue
            
        time_str = cols[0].get_text(strip=True).replace('\xa0', '').strip()
        time_match = re.match(r'^(\d{2,4})([A-Z]*)$', time_str)
        
        if time_match:
            t_str, orig_code = time_match.groups()
            t_str = t_str.zfill(4) if len(t_str) == 3 else (t_str + "00" if len(t_str) <= 2 else t_str)
            
            for idx, col in enumerate(cols[2:4]):
                is_dev = (idx == 1)
                cell_words = col.get_text(separator=' ', strip=True).replace('\xa0', ' ').split()
                
                for word in cell_words:
                    match = re.match(r'^([A-Z]{2})(\$?)$', word)
                    if match:
                        init, has_dollar = match.group(1), bool(match.group(2))
                        if init in ["OF", "OT"]: continue
                        
                        ot_hrs = ot_dict.get(init, {}).get('hrs', 0)
                        fos_hrs = ot_dict.get(init, {}).get('fos', 0)
                        
                        # Determine Code
                        if is_dev:
                            final_code = "T"
                        else:
                            final_code = orig_code
                            if has_dollar:
                                final_code = "$$" if orig_code == "A" or ot_hrs > 8 else "$"
                            elif not orig_code and (0 < ot_hrs < 8) and not has_dollar:
                                final_code = "$$"
                        
                        if fos_hrs > 0:
                            log_cb(f"FOS Detected -> {init} stays in {t_str} block, off-time math runs from {subtract_hours(t_str, fos_hrs)}")
                        
                        shifts.append({
                            'initials': init,
                            'start': t_str,
                            'orig_code': orig_code,
                            'code': final_code,
                            'ot': ot_hrs,
                            'fos': fos_hrs
                        })
                        
                        log_cb(f"Parsed -> {init} | Start: {t_str} | Code: {final_code or 'Standard (8-hr)'}{' (DEV)' if is_dev else ''}")
                            
    return shifts, leave_list, notes_list


# ==========================================
# 4. EXCEL GENERATION
# ==========================================
def build_excel_file(all_shifts, leave_list, notes_list, date_str, log_cb):
    morning_shifts = [s for s in all_shifts if s['start'] in ["2230", "2300", "2330", "2400"] or int(s['start']) < 1130]
    afternoon_shifts = [s for s in all_shifts if s not in morning_shifts]

    morning_shifts.sort(key=sort_chronological)
    afternoon_shifts.sort(key=sort_chronological)

    log_cb("Loading Excel template (This may take a moment)...")
    wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
    ws = wb["Projected Coverage"]
    
    def write_block(shifts_to_write, cols, label, start_row):
        row = start_row
        last_start = ""
        c_init, c_off, c_code, c_req = cols
        
        for shift in shifts_to_write:
            if last_start and shift['start'] != last_start:
                row += 1 
            
            off, req = calculate_times(shift['start'], shift['code'], shift['orig_code'], shift['ot'], shift['fos'])
            ws[f'{c_init}{row}'] = shift['initials']
            ws[f'{c_off}{row}'] = to_num(off)
            ws[f'{c_code}{row}'] = shift['code']
            ws[f'{c_req}{row}'] = to_num(req)
            
            last_start = shift['start']
            row += 1
        log_cb(f"Wrote {len(shifts_to_write)} {label} shifts to Excel.")

    write_block(morning_shifts, ('E', 'F', 'G', 'H'), 'Morning', 7)
    write_block(afternoon_shifts, ('J', 'K', 'L', 'M'), 'Afternoon', 9)

    log_cb(f"Writing {len(leave_list)} flagged requests to column AG...")
    for i, entry in enumerate(leave_list[:10]): # Max 10 rows
        ws[f'AG{9+i}'] = entry

    log_cb(f"Writing {len(notes_list)} shift notes to column AG...")
    for i, note in enumerate(notes_list[:10]): # Max 10 rows
        ws[f'AG{19+i}'] = note

    output_path = f"{OUTPUT_DIR}{date_str}.xlsm"
    log_cb(f"Saving workbook to: {output_path}")
    wb.save(output_path)
    return output_path


# ==========================================
# 5. CUSTOM UI COMPONENTS
# ==========================================
class MiniCalendar(tk.Toplevel):
    def __init__(self, parent, entry_widget):
        super().__init__(parent)
        self.entry_widget = entry_widget
        self.title("Select Date")
        self.geometry(f"+{parent.winfo_rootx() + 150}+{parent.winfo_rooty() + 200}")
        self.configure(bg=COLORS["bg_main"])
        self.resizable(False, False)
        self.attributes('-toolwindow', True)

        self.curr_year, self.curr_month = datetime.now().year, datetime.now().month
        
        self.header_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.header_frame.pack(fill=tk.X, pady=10, padx=10)
        
        tk.Button(self.header_frame, text="<", command=self.prev_month, bg=COLORS["btn_secondary"], fg=COLORS["fg_main"], relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT)
        self.month_lbl = tk.Label(self.header_frame, bg=COLORS["bg_main"], fg=COLORS["fg_main"], font=("Segoe UI", 10, "bold"))
        self.month_lbl.pack(side=tk.LEFT, expand=True)
        tk.Button(self.header_frame, text=">", command=self.next_month, bg=COLORS["btn_secondary"], fg=COLORS["fg_main"], relief=tk.FLAT, cursor="hand2").pack(side=tk.RIGHT)
        
        self.cal_frame = tk.Frame(self, bg=COLORS["bg_main"])
        self.cal_frame.pack(padx=10, pady=(0, 10))
        
        self.update_calendar()
        self.transient(parent)
        self.grab_set()

    def prev_month(self):
        self.curr_month = 12 if self.curr_month == 1 else self.curr_month - 1
        self.curr_year -= 1 if self.curr_month == 12 else 0
        self.update_calendar()

    def next_month(self):
        self.curr_month = 1 if self.curr_month == 12 else self.curr_month + 1
        self.curr_year += 1 if self.curr_month == 1 else 0
        self.update_calendar()

    def select_date(self, day):
        self.entry_widget.delete(0, tk.END)
        self.entry_widget.insert(0, f"{self.curr_month}-{day:02d}-{self.curr_year}")
        self.destroy()

    def update_calendar(self):
        for widget in self.cal_frame.winfo_children(): widget.destroy()
        self.month_lbl.config(text=f"{calendar.month_name[self.curr_month]} {self.curr_year}")
        
        for i, d in enumerate(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]):
            tk.Label(self.cal_frame, text=d, bg=COLORS["bg_main"], fg="#aaaaaa", font=("Segoe UI", 8)).grid(row=0, column=i, padx=2, pady=2)
            
        cal = calendar.Calendar(firstweekday=calendar.SUNDAY)
        for row_idx, week in enumerate(cal.monthdayscalendar(self.curr_year, self.curr_month)):
            for col_idx, day in enumerate(week):
                if day != 0:
                    btn = tk.Button(self.cal_frame, text=str(day), width=3, bg=COLORS["btn_secondary"], fg=COLORS["fg_main"], 
                                    activebackground=COLORS["btn_primary"], relief=tk.FLAT, cursor="hand2",
                                    command=lambda d=day: self.select_date(d))
                    btn.grid(row=row_idx+1, column=col_idx, padx=1, pady=1)

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, radius=15, width=150, height=40, bg_color=COLORS["btn_primary"], hover_color=COLORS["btn_hover"], font=("Segoe UI", 10, "bold"), **kwargs):
        super().__init__(parent, borderwidth=0, relief="flat", highlightthickness=0, bg=COLORS["bg_main"], width=width, height=height, cursor="hand2", **kwargs)
        self.command, self.bg_color, self.hover_color, self.radius, self.text_val, self.font = command, bg_color, hover_color, radius, text, font
        self.bind("<Configure>", self.draw)
        self.bind("<Enter>", lambda e: self.itemconfig(self.rect_id, fill=self.hover_color))
        self.bind("<Leave>", lambda e: self.itemconfig(self.rect_id, fill=self.bg_color))
        self.bind("<ButtonPress-1>", lambda e: self.move(self.text_id, 0, 1))
        self.bind("<ButtonRelease-1>", self.on_release)

    def draw(self, event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        r = self.radius
        points = [r,0, r,0, w-r,0, w-r,0, w,0, w,r, w,r, w,h-r, w,h-r, w,h, w-r,h, w-r,h, r,h, r,h, 0,h, 0,h-r, 0,h-r, 0,r, 0,r, 0,0]
        self.rect_id = self.create_polygon(points, fill=self.bg_color, smooth=True)
        self.text_id = self.create_text(w/2, h/2, text=self.text_val, fill="white", font=self.font)

    def on_release(self, event):
        self.move(self.text_id, 0, -1)
        if self.command: self.command()

    def config_text(self, new_text):
        self.itemconfig(self.text_id, text=new_text)


# ==========================================
# 6. MAIN APPLICATION CLASS
# ==========================================
class SpreaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TPA Watch Desk Spreader Creator - v1.0.1")
        
        # Load the custom icon for the title bar
        try:
            self.root.iconbitmap(APP_ICON)
        except Exception:
            pass # Fails safely if the network drive drops or the icon is moved
            
        # Default to smaller window, hiding the log on startup
        self.root.geometry("680x530")
        self.root.minsize(680, 530)
        self.root.configure(bg=COLORS["bg_main"])
        self.setup_styles()
        self.build_ui()
        self.log("System Initialized. Awaiting user input.")

    def setup_styles(self):
        style = ttk.Style()
        if 'clam' in style.theme_names(): style.theme_use('clam')
        style.configure('TFrame', background=COLORS["bg_main"])
        style.configure('TLabelframe', background=COLORS["bg_main"], foreground=COLORS["fg_main"])
        style.configure('TLabelframe.Label', background=COLORS["bg_main"], foreground=COLORS["fg_main"], font=("Segoe UI", 10, "bold"))
        style.configure('TLabel', background=COLORS["bg_main"], foreground=COLORS["fg_main"])

    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()

    def build_ui(self):
        main_container = ttk.Frame(self.root, padding=20)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(header_frame, text="Daily Spreader (semi)Automation", font=("Segoe UI", 18, "bold")).pack(anchor=tk.W)
        ttk.Label(header_frame, text="Extracts schedule data from WMT and builds the Excel template automatically.", foreground="#aaaaaa").pack(anchor=tk.W)

        # Step 1
        step1 = ttk.LabelFrame(main_container, text=" Step 1: Get the Schedule Data ", padding=(15, 10))
        step1.pack(fill=tk.X, pady=5)
        ttk.Label(step1, text="1. Click here to open the Worksheet view in your browser.", font=("Segoe UI", 9)).grid(row=0, column=0, sticky=tk.W, pady=5)
        RoundedButton(step1, "Launch WMT Scheduler", self.launch_wmt, radius=12, width=200, height=30, bg_color=COLORS["btn_help"], hover_color=COLORS["btn_help_hover"], font=("Segoe UI", 9, "bold")).grid(row=0, column=1, padx=15, pady=5)
        ttk.Label(step1, text="2. Navigate to your date, press Ctrl+S, and save as HTML.", font=("Segoe UI", 9), foreground="#aaaaaa").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        # Step 2
        step2 = ttk.LabelFrame(main_container, text=" Step 2: Build the Spreader ", padding=(15, 15))
        step2.pack(fill=tk.X, pady=10)
        
        ttk.Label(step2, text="Target Date:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=8, padx=5)
        self.date_entry = tk.Entry(step2, width=20, relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 10), bg=COLORS["bg_input"], fg=COLORS["fg_main"], insertbackground=COLORS["fg_main"])
        self.date_entry.grid(row=0, column=1, sticky=tk.W, pady=8, padx=5)
        tmrw = datetime.now() + timedelta(days=1)
        self.date_entry.insert(0, f"{tmrw.month}-{tmrw.strftime('%d-%Y')}")
        RoundedButton(step2, "Select Date", lambda: MiniCalendar(self.root, self.date_entry), radius=10, width=90, height=28, bg_color=COLORS["btn_secondary"], hover_color=COLORS["btn_sec_hover"], font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=10)

        ttk.Label(step2, text="Worksheet View HTML File:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=8, padx=5)
        self.file_entry = tk.Entry(step2, width=42, relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 10), bg=COLORS["bg_input"], fg=COLORS["fg_main"], insertbackground=COLORS["fg_main"])
        self.file_entry.grid(row=1, column=1, sticky=tk.W, pady=8, padx=5)
        RoundedButton(step2, "Browse...", self.browse_file, radius=15, width=90, height=28, bg_color=COLORS["btn_secondary"], hover_color=COLORS["btn_sec_hover"], font=("Segoe UI", 9, "bold")).grid(row=1, column=2, padx=10)

        # Actions
        actions = ttk.Frame(main_container)
        actions.pack(fill=tk.X, pady=5)
        RoundedButton(actions, "Generate Spreader", self.run_generation, radius=22, width=220, height=45, font=("Segoe UI", 11, "bold")).pack(pady=(10, 8))
        
        self.btn_help = RoundedButton(actions, "Help & Instructions", self.show_help, radius=12, width=160, height=28, bg_color=COLORS["bg_input"], hover_color=COLORS["btn_sec_hover"], font=("Segoe UI", 8, "bold"))
        self.btn_help.pack(pady=5)

        self.btn_toggle = RoundedButton(actions, "Show Debug Logger ⬇", self.toggle_log, radius=12, width=160, height=28, bg_color=COLORS["bg_input"], hover_color=COLORS["btn_sec_hover"], font=("Segoe UI", 8, "bold"))
        self.btn_toggle.pack(pady=(0, 5))

        # Log Frame (Created but not packed initially so it stays hidden)
        self.log_frame = ttk.LabelFrame(main_container, text=" System Log ", padding=(10, 10))
        scroll = ttk.Scrollbar(self.log_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(self.log_frame, height=12, wrap=tk.WORD, font=("Consolas", 10), yscrollcommand=scroll.set, borderwidth=1, relief=tk.SOLID, bg=COLORS["bg_log"], fg=COLORS["fg_log"], insertbackground=COLORS["fg_main"])
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scroll.config(command=self.log_text.yview)

    def toggle_log(self):
        if self.log_frame.winfo_ismapped():
            self.log_frame.pack_forget()
            self.btn_toggle.config_text("Show Debug Logger ⬇")
            self.root.geometry("680x530")
        else:
            self.log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            self.btn_toggle.config_text("Hide Debug Logger ⬆")
            self.root.geometry("680x780")

    def show_help(self):
        help_win = tk.Toplevel(self.root)
        help_win.title("Help & Instructions")
        help_win.geometry("550x650")
        help_win.configure(bg=COLORS["bg_main"])
        
        try:
            help_win.iconbitmap(APP_ICON)
        except Exception:
            pass

        title = tk.Label(help_win, text="Spreader Creator - User Guide", bg=COLORS["bg_main"], fg=COLORS["fg_main"], font=("Segoe UI", 14, "bold"))
        title.pack(pady=(15, 10))

        text_frame = tk.Frame(help_win, bg=COLORS["bg_main"])
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        scroll = tk.Scrollbar(text_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        help_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scroll.set, bg=COLORS["bg_input"], fg=COLORS["fg_main"], font=("Segoe UI", 10), relief=tk.FLAT, padx=10, pady=10)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=help_text.yview)

        instructions = (
            "1. DOWNLOADING WMT DATA\n"
            "----------------------------------------\n"
            "• Click 'Launch WMT Scheduler' to open Chrome.\n"
            "• Navigate to the Worksheet View for your target date.\n"
            "• Press Ctrl+S to save the page.\n"
            "• CRITICAL: Change 'Save as type' to 'Webpage, HTML Only'.\n"
            "• Save it to the network Downloads folder.\n\n"
            
            "2. GENERATING THE SPREADER\n"
            "----------------------------------------\n"
            "• Select the Target Date using the built-in calendar.\n"
            "• Click 'Browse' and select the HTML file you just saved.\n"
            "• Click 'Generate Spreader'.\n\n"
            
            "3. HOW IT WORKS BEHIND THE SCENES\n"
            "----------------------------------------\n"
            "• DEV Shifts: Automatically marked with a 'T' code. Their End Time is placed in the REQ column, and OFF is left blank so they don't count towards staffing numbers.\n"
            "• Overtime ($ and $$): Automatically recognized. The script knows standard 8-hour OT ($) vs extended 10-hour OT ($$).\n"
            "• FOS (Early Starts): Automatically adjusts the start time for the math, sets an 8-hour OFF time from that early start, and drops the final extended end time into the REQ column.\n"
            "• Leave Requests: Automatically flags unapproved leave (both full and partial shifts with times) to column AG. Ignores approved leave, shift changes, and shift swaps.\n"
            "• Shift Notes: Imports notes cleanly while filtering out automated WMT system notes (like '65 - OT OJT').\n\n"
            
            "4. TROUBLESHOOTING\n"
            "----------------------------------------\n"
            "• Error: 'No valid shifts found' -> The WMT file was likely saved as 'Webpage, Complete' instead of 'HTML Only'. Delete it, re-save correctly, and try again.\n"
            "• File won't open automatically -> The spreader was created successfully, but Excel might be busy. Navigate to the Watch Desk Documents folder to open it manually."
        )

        help_text.insert(tk.END, instructions)
        help_text.config(state=tk.DISABLED)

    def launch_wmt(self):
        try:
            paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
            chrome_path = next((p for p in paths if os.path.exists(p)), None)
            
            if chrome_path:
                webbrowser.register('chrome_explicit', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome_explicit').open_new_tab(WMT_URL)
            else:
                webbrowser.open_new_tab(WMT_URL)
                
            messagebox.showinfo("Step 1", "WMT Scheduler should be opening in Chrome!\n\n1. Navigate to target date.\n2. Press 'Ctrl + S'.\n3. Change 'Save as type' to 'Webpage, HTML Only'.\n4. Save to network Downloads folder.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open browser:\n{str(e)}")

    def browse_file(self):
        filepath = filedialog.askopenfilename(initialdir=DOWNLOADS_DIR, title="Select WMT HTML File", filetypes=[("HTML Files", "*.html *.htm")])
        if filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filepath)

    def run_generation(self):
        self.log("--- INITIATING SPREADER GENERATION ---")
        filepath = self.file_entry.get()
        if not filepath:
            self.log("ERROR: No HTML file path provided.")
            messagebox.showwarning("Missing File", "Please select a WMT HTML file first.")
            return

        try:
            shifts, leaves, notes = parse_html_data(filepath, self.log)
            if not shifts:
                self.log("PROCESS HALTED: No valid shifts found.")
                messagebox.showerror("No Data", "Could not find any shifts. Did you save the WMT page correctly?")
                return
                
            out_path = build_excel_file(shifts, leaves, notes, self.date_entry.get(), self.log)
            self.log("--- SUCCESS! SPREADER COMPLETE ---")
            
            if messagebox.askyesno("Success", f"Spreader created successfully!\n\nSaved as: {self.date_entry.get()}.xlsm\n\nWould you like to open it now?"):
                try: os.startfile(out_path)
                except Exception as e:
                    self.log(f"Native open failed: {e}")
                    messagebox.showerror("Error", "Could not open automatically. File is in the Watch Desk folder.")

        except Exception as e:
            self.log(f"CRITICAL ERROR: {str(e)}")
            self.log(traceback.format_exc())
            messagebox.showerror("Error", "A critical error occurred. Check the debug log for details.")


if __name__ == "__main__":
    root = tk.Tk()
    app = SpreaderApp(root)
    root.mainloop()