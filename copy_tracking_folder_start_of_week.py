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
