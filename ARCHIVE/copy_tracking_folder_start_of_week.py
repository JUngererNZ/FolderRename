"""
copy_tracking_folder_start_of_week.py

Variant of copy_tracking_folder.py intended for start-of-week runs.
Functionally identical except for two differences: uses timedelta(days=0)
as a placeholder for date offsetting, and omits BARTRAC - SURYA MINES
from the prefix rename map.

Copies the most recent date-named tracking subfolder into a new folder
named with today's date, then renames all .xlsx files using today's date suffix.

Usage:
    python copy_tracking_folder_start_of_week.py

    No arguments. base_dir and prefix mappings are hardcoded below.

Difference from copy_tracking_folder.py:
    - BARTRAC - SURYA MINES is absent from the prefix map (intentional for
      start-of-week, or an omission — verify with Jason)
    - timedelta(days=0) is a no-op but left as a hook to offset the target
      date if needed (e.g. days=1 to pre-create tomorrow's folder)

Limitations:
    - base_dir is hardcoded — update when the month folder changes
    - timedelta(days=0) makes this identical to copy_tracking_folder.py
      until the offset is changed — consider whether this variant is still needed
    - Will fail if today's folder already exists (shutil.copytree raises FileExistsError)
    - MUMI and SURYA MINES absent from prefix mappings — verify if intentional
"""

import os
import shutil
from datetime import datetime, timedelta

# script updated to look at production tracking folder. 
base_dir = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\TRACKING\APRIL 2026"
next_day = datetime.today() + timedelta(days=0)
today_str = next_day.strftime("%d-%m-%Y")
target_path = os.path.join(base_dir, today_str)

# Find latest date folder
folders = []
for d in os.listdir(base_dir):
    full = os.path.join(base_dir, d)
    if os.path.isdir(full):
        try:
            folders.append((datetime.strptime(d, "%d-%m-%Y"), full))
        except ValueError:
            pass

if not folders:
    print("No date-named subfolders found.")
    exit(1)

latest_folder = sorted(folders)[-1][1]
print(f"Latest folder: {latest_folder}")
print(f"Copying to:    {target_path}")

shutil.copytree(latest_folder, target_path)

# Rename files
prefixes = {
    "BARTRAC - CONGO TRACKING":                   f"BARTRAC - CONGO TRACKING {today_str}.xlsx",
    "BARTRAC - KCC TRACKING AS OF":               f"BARTRAC - KCC TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - TRACKING - FML BONDED FACILITY -": f"BARTRAC - TRACKING - FML BONDED FACILITY - {today_str}.xlsx",
    "FML-KANU - ALLAN - TRACKING AS OF":          f"FML-KANU - ALLAN - TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - KAMOA TRACKING AS OF":             f"BARTRAC - KAMOA TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - ERG TRACKING":                     f"BARTRAC - ERG TRACKING {today_str}.xlsx",
}

for fname in os.listdir(target_path):
    if not fname.endswith(".xlsx"):
        continue
    for prefix, new_name in prefixes.items():
        if fname.startswith(prefix):
            os.rename(
                os.path.join(target_path, fname),
                os.path.join(target_path, new_name)
            )
            print(f"Renamed: {fname} -> {new_name}")
            break

print(f"\nDone. Folder and files updated to {today_str}")
