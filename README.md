# Tracking Folder Copy Script

This Python script automates the creation of daily tracking folders and files for FML Freight Solutions.

## Purpose

The script is designed to:
1. Create a new folder with today's date in the format `DD-MM-YYYY`
2. Copy all files from the latest existing date folder
3. Rename specific tracking files to include the current date
4. Maintain the folder structure including the `COMPLETE` subdirectory

## Usage

Run the script from the command line:
```bash
python copy_tracking_folder.py
```

## Current Configuration

The script is currently configured to:
- **Base Directory**: `C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\FolderRename\MARCH 2026`
- **Date Format**: `DD-MM-YYYY` (e.g., `06-03-2026`)
- **Files Created**: 
  - `BARTRAC - CONGO TRACKING {date}.xlsx`
  - `BARTRAC - KCC TRACKING AS OF {date}.xlsx`
  - `BARTRAC - TRACKING - FML BONDED FACILITY - {date}.xlsx`
  - `FML-KANU - ALLAN - TRACKING AS OF {date}.xlsx`

## Folder Structure

```
MARCH 2026/
├── 02-03-2026/
│   ├── BARTRAC - CONGO TRACKING 02-03-2026.xlsx
│   ├── BARTRAC - KCC TRACKING AS OF 02-03-2026.xlsx
│   ├── BARTRAC - TRACKING - FML BONDED FACILITY - 02-03-2026.xlsx
│   ├── FML-KANU - ALLAN - TRACKING AS OF 02-03-2026.xlsx
│   └── COMPLETE/
├── 03-03-2026/
│   ├── BARTRAC - CONGO TRACKING 03-03-2026.xlsx
│   ├── BARTRAC - KCC TRACKING AS OF 03-03-2026.xlsx
│   ├── BARTRAC - TRACKING - FML BONDED FACILITY - 03-03-2026.xlsx
│   ├── FML-KANU - ALLAN - TRACKING AS OF 03-03-2026.xlsx
│   └── COMPLETE/
└── ...
```

## Monday Configuration

To create folders and files for a specific Monday (e.g., Monday 09-03-2026), modify the script as follows:

### Current Code (Lines 4-7):
```python
base_dir = r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\FolderRename\MARCH 2026"
today_str = datetime.today().strftime("%d-%m-%Y")
target_path = os.path.join(base_dir, today_str)
```

### Modified Code for Monday 09-03-2026:
```python
base_dir = r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\FolderRename\MARCH 2026"

# Use a specific date for Monday 09-03-2026
target_date = datetime(2026, 3, 9)
today_str = target_date.strftime("%d-%m-%Y")
target_path = os.path.join(base_dir, today_str)
```

### Steps to Implement Monday Configuration:

1. **Locate the date configuration section** (lines 4-7 in the script)
2. **Replace the current code** with the modified version above
3. **Change the date** in `datetime(2026, 3, 9)` to your desired Monday:
   - Format: `datetime(year, month, day)`
   - Example for Monday 16-03-2026: `datetime(2026, 3, 16)`

### Reverting to Current Date:

To revert back to using today's date, replace the Monday configuration with:
```python
base_dir = r"C:\Users\Jason\OneDrive - FML Freight Solutions\FML-PROJECTS\FolderRename\MARCH 2026"
today_str = datetime.today().strftime("%d-%m-%Y")
target_path = os.path.join(base_dir, today_str)
```

## Notes

- The script automatically finds the latest date folder to copy from
- All files are copied including the `COMPLETE` subdirectory
- Only `.xlsx` files are processed for renaming
- The script will exit if no date-named subfolders are found
- Always verify the created folder and files after running the script