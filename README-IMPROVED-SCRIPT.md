# Enhanced Excel Column Management Script

## Overview

The `Insert-and-hide-a-column-improved.py` script is an enhanced version that automatically:

1. **Finds the last column** containing "COMMENTS [date]" header
2. **Copies all data** from that column (header + all rows)
3. **Inserts a new column** after the last column
4. **Pastes the copied data** into the new column
5. **Hides the previous column** (the one that was copied from)
6. **Adds a new header** with today's date


## Usage

### Interactive Mode (Recommended - Opens File Picker)
```bash
python Insert-and-hide-a-column-improved.py
```
This will open a file dialog where you can browse and select your Excel file.

### Command Line Mode
```bash
python Insert-and-hide-a-column-improved.py "path/to/your/file.xlsx"
```

### With Backup (Recommended)
```bash
python Insert-and-hide-a-column-improved.py "path/to/your/file.xlsx" --backup
```

### With Custom Sheet Name
```bash
python Insert-and-hide-a-column-improved.py "path/to/your/file.xlsx" --sheet "Sheet1"
```

### With Custom Date Format
```bash
python Insert-and-hide-a-column-improved.py "path/to/your/file.xlsx" --date-format "%Y-%m-%d"
```

## Command Line Arguments

- `file` (required): Path to the Excel .xlsx file
- `-s, --sheet` (optional): Sheet name (default: active sheet)
- `-b, --backup` (optional): Create backup before modifying
- `-d, --date-format` (optional): Date format for header (default: %d-%m-%Y)

## Example Output

```
✓ Backup created: file.backup.xlsx
✓ Working on sheet: ENROUTE SITE
✓ Found last column: FH (column 164)
✓ Copied 8 rows from column FH
✓ Inserted new column at FI (column 165)
✓ Pasted data into new column FI
✓ Hidden previous column FH
✓ Set FI1 = 'COMMENTS 11-03-2026'
✓ Saved: file.xlsx

✅ Complete! New column inserted with copied data and previous column hidden.
```

## Key Improvements Over Original Script

1. **Automatic Column Detection**: No need to specify column letters manually
2. **Flexible Header Location**: Searches multiple rows for COMMENTS headers
3. **Data Preservation**: Copies all existing data from the last column
4. **Dynamic Column Insertion**: Works with any number of columns
5. **Robust Error Handling**: Gracefully handles various Excel file structures

## Requirements

- Python 3.x
- openpyxl library: `pip install openpyxl`
- pandas library (for analysis): `pip install pandas`

## File Structure

```
FolderRename/
├── Insert-and-hide-a-column-improved.py    # Enhanced script
├── README-IMPROVED-SCRIPT.md              # This documentation
├── MARCH 2026/
│   └── 09-03-2026/
│       ├── BARTRAC - KCC TRACKING AS OF 09-03-2026.xlsx
│       └── BARTRAC - KCC TRACKING AS OF 09-03-2026.backup.xlsx
└── ...
```

## Testing

The script has been tested on the KCC tracking file and successfully:

1. ✅ Found the last COMMENTS column (FH - column 164)
2. ✅ Copied all data from that column
3. ✅ Inserted a new column (FI - column 165)
4. ✅ Pasted the copied data
5. ✅ Hidden the previous column
6. ✅ Added new header: "COMMENTS 11-03-2026"


## Notes

- The script creates a backup file with `.backup.xlsx` extension when using the `--backup` flag
- Hidden columns are not visible in Excel but the data is preserved
- The script works with any Excel file that has COMMENTS columns with date headers
- Column detection searches the first 10 rows to find COMMENTS headers