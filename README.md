# Daily Spreader (semi)Automation

A Python-based GUI tool designed for TPA Operations Supervisors to streamline the creation of the Daily Spreader. This tool extracts shift schedules, overtime, leave requests, and shift notes directly from the WMT Scheduler "Worksheet View" and automatically populates the official Excel template.

## First-Time Setup (Important)

Before using the tool for the first time on any workstation, you must install the required background engines. You do not need IT Admin privileges to do this.

1. Navigate to the shared Watch Desk folder.
2. Double-click **`IC-Tools\Spreader Creator\Install_Python.bat`**. This may need to be ran Twice see #4
3. **What this does:**
   * It checks if Python is installed on your PC. 
   * If Python is missing, it silently installs it directly from the network (`IC-Tools\python-manager-26.3.msix`).
   * It installs the two required data engines (`openpyxl` for Excel and `beautifulsoup4` for HTML).
4. **Note:** If the script has to install Python for the first time, it will ask you to close the window and double-click the script **one more time** to finish the setup.
5. Once the script says `[SUCCESS]`, you are ready to go.

## Daily Usage Instructions

To open the tool, double-click **`IC-Tools\Spreader Creator\spreader_creator.py`**.

### Step 1: Get the Schedule Data
1. Click the **Launch WMT Scheduler** button inside the app. This will open Google Chrome to the WMT Worksheet View.
2. Navigate to your target date.
3. Press `Ctrl + S` on your keyboard.
4. Change "Save as type" to **Webpage, HTML Only**.
5. Save the file to the network **Downloads** folder (`\Watch Desk Documents\Downloads`).

### Step 2: Build the Spreader
1. **Target Date:** The app automatically defaults to tomorrow's date. If you are running it for a different day, click the **Select Date** button to pick the correct day from the calendar.
2. **Worksheet View HTML File:** Click **Browse...** and select the HTML file you just downloaded.
3. Click **Generate Spreader**.
4. The tool will parse the data, flag FOS adjustments, format the shift notes, and build the Excel file. 
5. When complete, a prompt will ask if you want to open the newly created Spreader to review it.

## Output Details

* **File Naming:** The completed workbook is automatically saved to the main Watch Desk Documents folder as `M-DD-YYYY.xlsm`.
* **Troubleshooting:** If someone's shift populates incorrectly or a note looks strange, click the **Show Debug Logger** button in the app. The logger provides a highly detailed breakdown of exactly how it calculated every shift and OT code so you can easily trace the error.
