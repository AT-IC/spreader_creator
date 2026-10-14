import calendar
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox
import webbrowser

import customtkinter as ctk

from config import APP_ICON, CONFIG, DOWNLOADS_DIR, TEMPLATE_PATH, OUTPUT_DIR, WMT_URL
from excel_writer import build_excel_file
from parser import parse_html_data, preview_text


class MiniCalendar(ctk.CTkToplevel):
    def __init__(self, parent, entry_widget):
        super().__init__(parent)
        self.entry_widget = entry_widget
        self.title("Select Date")
        self.geometry(f"300x330+{parent.winfo_rootx() + 150}+{parent.winfo_rooty() + 150}")
        self.configure(fg_color=CONFIG["bg_main"])
        self.resizable(False, False)
        self.curr_year, self.curr_month = datetime.now().year, datetime.now().month
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=18, pady=(18, 10))
        for text, command, side in (("‹", self.prev_month, "left"), ("›", self.next_month, "right")):
            ctk.CTkButton(self.header_frame, text=text, width=36, height=32, corner_radius=12,
                          fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"],
                          command=command).pack(side=side)
        self.month_lbl = ctk.CTkLabel(self.header_frame, text="", text_color=CONFIG["fg_main"],
                                      font=(CONFIG["font_family"], 16, "bold"))
        self.month_lbl.pack(side="left", expand=True)
        self.cal_frame = ctk.CTkFrame(self, fg_color=CONFIG["bg_surface"], corner_radius=18)
        self.cal_frame.pack(padx=18, pady=(0, 18), fill="both", expand=True)
        self.update_calendar()
        self.transient(parent)
        self.grab_set()

    def prev_month(self):
        self.curr_month = 12 if self.curr_month == 1 else self.curr_month - 1
        self.curr_year -= self.curr_month == 12
        self.update_calendar()

    def next_month(self):
        self.curr_month = 1 if self.curr_month == 12 else self.curr_month + 1
        self.curr_year += self.curr_month == 1
        self.update_calendar()

    def select_date(self, day):
        self.entry_widget.delete(0, tk.END)
        self.entry_widget.insert(0, f"{self.curr_month}-{day:02d}-{self.curr_year}")
        self.destroy()

    def update_calendar(self):
        for widget in self.cal_frame.winfo_children():
            widget.destroy()
        self.month_lbl.configure(text=f"{calendar.month_name[self.curr_month]} {self.curr_year}")
        for i, day_name in enumerate(("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")):
            ctk.CTkLabel(self.cal_frame, text=day_name, text_color=CONFIG["fg_muted"],
                         font=(CONFIG["font_family"], 12, "bold")).grid(row=0, column=i, padx=3, pady=(12, 5))
        for row_idx, week in enumerate(calendar.Calendar(firstweekday=calendar.SUNDAY).monthdayscalendar(self.curr_year, self.curr_month)):
            for col_idx, day in enumerate(week):
                if day:
                    ctk.CTkButton(self.cal_frame, text=str(day), width=29, height=27, corner_radius=10,
                                  fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["accent"],
                                  command=lambda selected=day: self.select_date(selected)).grid(row=row_idx + 1, column=col_idx, padx=3, pady=3)


class SpreaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TPA Watch Desk Spreader Creator - v1.0.1")
        try:
            self.root.iconbitmap(APP_ICON)
        except Exception:
            pass
        self.root.geometry("760x690")
        self.root.minsize(700, 620)
        self.root.configure(fg_color=CONFIG["bg_main"])
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.build_ui()
        self.log("System Initialized. Awaiting user input.")

    def log(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")

    def build_ui(self):
        container = ctk.CTkFrame(self.root, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=30, pady=26)
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(header, text="Daily Spreader", text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(header, text="Semi-automation for the TPA Watch Desk", text_color=CONFIG["accent"], font=(CONFIG["font_family"], 15, "bold")).pack(anchor="w")
        ctk.CTkLabel(header, text="Pull schedule data from WMT and build the Excel spreader in a few clicks.", text_color=CONFIG["fg_muted"], font=(CONFIG["font_family"], 14)).pack(anchor="w", pady=(7, 0))
        step1 = ctk.CTkFrame(container, fg_color=CONFIG["bg_surface"], corner_radius=CONFIG["radius_card"])
        step1.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(step1, text="01  GET THE SCHEDULE DATA", text_color=CONFIG["accent"], font=(CONFIG["font_family"], 14, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=24, pady=(20, 4))
        ctk.CTkLabel(step1, text="Open the WMT Worksheet View, then save the target date as HTML.", text_color=CONFIG["fg_muted"], font=(CONFIG["font_family"], 13)).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 18))
        ctk.CTkButton(step1, text="Launch WMT Scheduler", command=self.launch_wmt, width=210, height=42, fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"], font=(CONFIG["font_family"], 13, "bold")).grid(row=1, column=1, padx=(10, 24), pady=(0, 18))
        step2 = ctk.CTkFrame(container, fg_color=CONFIG["bg_surface"], corner_radius=CONFIG["radius_card"])
        step2.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(step2, text="02  BUILD THE SPREADER", text_color=CONFIG["accent"], font=(CONFIG["font_family"], 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=24, pady=(20, 15))
        ctk.CTkLabel(step2, text="Target date", text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 13, "bold")).grid(row=1, column=0, sticky="w", padx=(24, 12), pady=8)
        self.date_entry = ctk.CTkEntry(step2, height=40, fg_color=CONFIG["bg_input"], border_color=CONFIG["border"], text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 13))
        self.date_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=8)
        tomorrow = datetime.now() + timedelta(days=1)
        self.date_entry.insert(0, f"{tomorrow.month}-{tomorrow.day:02d}-{tomorrow.year}")
        ctk.CTkButton(step2, text="Select date", command=lambda: MiniCalendar(self.root, self.date_entry), width=105, height=40, fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"], font=(CONFIG["font_family"], 12, "bold")).grid(row=1, column=2, padx=(8, 24), pady=8)
        ctk.CTkLabel(step2, text="Worksheet HTML file", text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 13, "bold")).grid(row=2, column=0, sticky="w", padx=(24, 12), pady=(8, 20))
        self.file_entry = ctk.CTkEntry(step2, height=40, fg_color=CONFIG["bg_input"], border_color=CONFIG["border"], text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 12))
        self.file_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=(8, 20))
        ctk.CTkButton(step2, text="Browse…", command=self.browse_file, width=105, height=40, fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"], font=(CONFIG["font_family"], 12, "bold")).grid(row=2, column=2, padx=(8, 24), pady=(8, 20))
        step2.grid_columnconfigure(1, weight=1)
        self.dry_run_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(container, text="Preview / validate only (do not save Excel)", variable=self.dry_run_var, font=(CONFIG["font_family"], 12)).pack(anchor="w", pady=(0, 8))
        self.btn_generate = ctk.CTkButton(container, text="Generate Spreader", command=self.run_generation, height=54, fg_color=CONFIG["accent"], hover_color=CONFIG["accent_hover"], text_color=CONFIG["fg_on_accent"], font=(CONFIG["font_family"], 16, "bold"))
        self.btn_generate.pack(fill="x", pady=(3, 14))
        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.pack(fill="x")
        self.btn_help = ctk.CTkButton(actions, text="Help & Instructions", command=self.show_help, height=34, fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"], font=(CONFIG["font_family"], 12, "bold"))
        self.btn_help.pack(side="left", padx=(0, 8))
        self.btn_toggle = ctk.CTkButton(actions, text="Show Debug Logger ↓", command=self.toggle_log, height=34, fg_color=CONFIG["btn_secondary"], hover_color=CONFIG["btn_sec_hover"], font=(CONFIG["font_family"], 12, "bold"))
        self.btn_toggle.pack(side="left")
        self.log_frame = ctk.CTkFrame(container, fg_color=CONFIG["bg_surface"], corner_radius=CONFIG["radius_card"])
        ctk.CTkLabel(self.log_frame, text="SYSTEM LOG", text_color=CONFIG["accent"], font=(CONFIG["font_family"], 13, "bold")).pack(anchor="w", padx=18, pady=(15, 8))
        self.log_text = ctk.CTkTextbox(self.log_frame, height=220, fg_color=CONFIG["bg_log"], text_color=CONFIG["fg_log"], font=("Consolas", 12), wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=14, pady=(0, 14))

    def toggle_log(self):
        if self.log_frame.winfo_ismapped():
            self.log_frame.pack_forget()
            self.btn_toggle.configure(text="Show Debug Logger ↓")
            self.root.geometry("760x690")
        else:
            self.log_frame.pack(fill="both", expand=True, pady=(14, 0))
            self.btn_toggle.configure(text="Hide Debug Logger ↑")
            self.root.geometry("760x930")

    def show_help(self):
        help_win = ctk.CTkToplevel(self.root)
        help_win.title("Help & Instructions")
        help_win.geometry("550x650")
        help_win.configure(fg_color=CONFIG["bg_main"])
        try:
            help_win.iconbitmap(APP_ICON)
        except Exception:
            pass
        title = ctk.CTkLabel(help_win, text="Spreader Creator - User Guide", text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 20, "bold"))
        title.pack(pady=(15, 10))
        help_text = ctk.CTkTextbox(help_win, wrap="word", fg_color=CONFIG["bg_input"], text_color=CONFIG["fg_main"], font=(CONFIG["font_family"], 13))
        help_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        help_text.insert("end", "1. Download the WMT Worksheet View as Webpage, HTML Only.\n\n"
                              "2. Select the target date and HTML file.\n\n"
                              "3. Use Preview / validate only to inspect parsed data without saving Excel.\n\n"
                              "4. Clear the preview checkbox to create the workbook.\n\n"
                              "The Debug Logger shows parser and workbook status messages.")
        help_text.configure(state="disabled")
        help_win.transient(self.root)

    def launch_wmt(self):
        try:
            webbrowser.open_new_tab(WMT_URL)
            messagebox.showinfo("Step 1", "Save the target WMT page as Webpage, HTML Only.")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open browser:\n{exc}")

    def browse_file(self):
        filepath = filedialog.askopenfilename(initialdir=DOWNLOADS_DIR, title="Select WMT HTML File", filetypes=[("HTML Files", "*.html *.htm")])
        if filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filepath)

    def run_generation(self):
        filepath = self.file_entry.get().strip()
        if not filepath:
            messagebox.showwarning("Missing File", "Please select a WMT HTML file first.")
            return
        try:
            data = parse_html_data(filepath, self.log)
            if not data.shifts:
                messagebox.showerror("No Data", "Could not find any shifts in the HTML file.")
                return
            if self.dry_run_var.get():
                messagebox.showinfo("Preview", preview_text(data))
                build_excel_file(data, self.date_entry.get(), TEMPLATE_PATH, OUTPUT_DIR, self.log, dry_run=True)
                return
            output_path = build_excel_file(data, self.date_entry.get(), TEMPLATE_PATH, OUTPUT_DIR, self.log)
            if messagebox.askyesno("Success", f"Spreader created:\n\n{output_path}\n\nOpen it now?"):
                try:
                    os.startfile(output_path)
                except (AttributeError, OSError) as exc:
                    messagebox.showerror("Open failed", str(exc))
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Error", str(exc))


def run_app():
    root = ctk.CTk()
    SpreaderApp(root)
    root.mainloop()
