"""
copy_tracking_folder.py

Step 1 of the consolidated BARTRAC weekly workflow. Copies the most recent
date-named tracking subfolder into a new folder named with today's date, then
renames all .xlsx files inside it using today's date suffix.

Finds the latest folder by parsing subfolder names against the DD-MM-YYYY format,
copies the entire folder tree to a new sibling folder named with today's date,
then matches each .xlsx file against a prefix map and renames it accordingly.

Usage:
    python copy_tracking_folder.py

    No arguments. base_dir and prefix mappings are hardcoded below.

Folder structure:
    TRACKING/APRIL 2026/
        05-05-2026/                  ← created by this script
            BARTRAC - KCC TRACKING AS OF 05-05-2026.xlsx
            BARTRAC - KAMOA TRACKING AS OF 05-05-2026.xlsx
            BARTRAC - SURYA MINES 05-05-2026.xlsx
            ... (all clients)

Prefix mappings (source → renamed):
    BARTRAC - CONGO TRACKING                   → BARTRAC - CONGO TRACKING {date}.xlsx
    BARTRAC - KAMOA TRACKING AS OF             → BARTRAC - KAMOA TRACKING AS OF {date}.xlsx
    BARTRAC - KCC TRACKING AS OF               → BARTRAC - KCC TRACKING AS OF {date}.xlsx
    BARTRAC - TRACKING - FML BONDED FACILITY - → BARTRAC - TRACKING - FML BONDED FACILITY - {date}.xlsx
    FML-KANU - ALLAN - TRACKING AS OF          → FML-KANU - ALLAN - TRACKING AS OF {date}.xlsx
    BARTRAC - ERG TRACKING                     → BARTRAC - ERG TRACKING {date}.xlsx
    BARTRAC - SURYA MINES                      → BARTRAC - SURYA MINES {date}.xlsx

Limitations:
    - base_dir is hardcoded — update when the month folder changes (e.g. APRIL → MAY 2026)
    - No --backup flag; source folder is untouched but destination is not recoverable if run twice
    - Will fail if today's folder already exists (shutil.copytree raises FileExistsError)
    - MUMI is absent from prefix mappings — verify if intentional
"""

import os
import shutil
from datetime import datetime

# script updated to look at production tracking folder. 
base_dir = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\TRACKING\APRIL 2026"
# base_dir = r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\FolderRename\MARCH 2026"
today_str = datetime.today().strftime("%d-%m-%Y")
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
    "BARTRAC - KAMOA TRACKING AS OF":             f"BARTRAC - KAMOA TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - KCC TRACKING AS OF":               f"BARTRAC - KCC TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - TRACKING - FML BONDED FACILITY -": f"BARTRAC - TRACKING - FML BONDED FACILITY - {today_str}.xlsx",
    "FML-KANU - ALLAN - TRACKING AS OF":          f"FML-KANU - ALLAN - TRACKING AS OF {today_str}.xlsx",
    "BARTRAC - ERG TRACKING":                     f"BARTRAC - ERG TRACKING {today_str}.xlsx",
    "BARTRAC - SURYA MINES":                     f"BARTRAC - SURYA MINES {today_str}.xlsx",
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
