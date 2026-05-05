"""
copy-insert-hide-column.py

Automates the copy → insert → hide column operation on BARTRAC Excel tracking workbooks.

Scans a specified row for the rightmost column matching the pattern 'COMMENTS DD-MM-YYYY',
copies its data, inserts a hidden duplicate column to its left, then saves the file.
This preserves the previous week's comments before a new dated column is added.

Usage:
    python copy-insert-hide-column.py <file.xlsx> [--backup]

Notes:
    - SEARCH_ROW defaults to 6 — adjust per client (see tracking_workflow_config.json)
    - Targets the active sheet; ensure the correct sheet is active before running
    - --backup creates a .bak.xlsx copy before any modifications
"""


import xlwings as xw
import argparse
import re
from pathlib import Path
import shutil

# Configuration constants
COMMENTS_PATTERN = r'COMMENTS\s+\d{2}-\d{2}-\d{4}'
SEARCH_ROW = 6

def find_last_comments_column(worksheet, search_row=SEARCH_ROW):
    """
    Find the rightmost column containing 'COMMENTS DD-MM-YYYY' pattern.
    
    Args:
        worksheet: xlwings worksheet object
        search_row: Row number to search for the pattern (default: 6)
    
    Returns:
        int: Column number of the last matching column, or None if not found
    """
    pattern = re.compile(COMMENTS_PATTERN, re.IGNORECASE)
    last_col = None
    max_col = worksheet.cells.last_cell.column
    
    print(f"Searching for pattern '{COMMENTS_PATTERN}' in row {search_row}...")
    
    for col in range(1, max_col + 1):
        cell_value = worksheet.range((search_row, col)).value
        if cell_value and pattern.search(str(cell_value)):
            last_col = col
            print(f"  Found match in column {xw.utils.col_name(col)}: '{cell_value}'")
    
    return last_col

def copy_column_data(worksheet, column, last_row):
    """
    Copy data from a specific column.
    
    Args:
        worksheet: xlwings worksheet object
        column: Column number to copy from
        last_row: Last row number to include in the copy
    
    Returns:
        list: Copied data from the column
    """
    return worksheet.range((1, column), (last_row, column)).value

def insert_and_hide_column(worksheet, column, copied_data, last_row):
    """
    Insert a new column and hide it with copied data.
    
    Args:
        worksheet: xlwings worksheet object
        column: Column number where to insert (1-based)
        copied_data: Data to paste into the new column
        last_row: Last row number for the paste range
    """
    # Insert new column before the specified column
    insert_range = worksheet.range((1, column))
    insert_range.insert(shift='right')
    
    # Paste copied data into the new column (now at position column)
    paste_range = worksheet.range((1, column), (last_row, column))
    paste_range.value = copied_data
    
    # Hide the newly inserted column
    worksheet.range((1, column)).api.EntireColumn.Hidden = True
    
    return column

def main():
    parser = argparse.ArgumentParser(description="Copy last column with data, insert new column before it, hide the new column.")
    parser.add_argument('file', nargs='?', help='Path to Excel file')
    parser.add_argument('--backup', action='store_true', help='Create backup before modifying')
    args = parser.parse_args()

    if args.file:
        file_path = Path(args.file)
    else:
        file_path_str = input("Enter the path to the Excel file: ")
        if not file_path_str:
            print("No file path provided")
            return
        file_path = Path(file_path_str)

    if not file_path.exists():
        print(f"File not found: {file_path}")
        return

    # Create backup if requested
    if args.backup:
        backup_path = file_path.with_suffix('.bak.xlsx')
        shutil.copy(file_path, backup_path)
        print(f"Backup created: {backup_path}")

    try:
        # Open workbook with xlwings
        wb = xw.Book(file_path)
        ws = wb.sheets.active
        
        print(f"Active sheet: {ws.name}")
        
        # Find the last column with "COMMENTS DD-MM-YYYY" pattern
        last_col = find_last_comments_column(ws, SEARCH_ROW)
        
        if last_col is None:
            print(f"No column found with pattern '{COMMENTS_PATTERN}' in row {SEARCH_ROW}")
            wb.close()
            return

        column_name = xw.utils.col_name(last_col)
        cell_value = ws.range((SEARCH_ROW, last_col)).value
        print(f"Found target column: {column_name} with value: '{cell_value}'")

        # Get the actual used range to determine the real last row
        used_range = ws.used_range
        last_row = used_range.rows.count
        print(f"Data range: rows 1 to {last_row}")

        # Copy data from the target column
        print("Copying data from target column...")
        copied_data = copy_column_data(ws, last_col, last_row)

        # Insert new column before the target column and hide it
        print(f"Inserting new column before column {column_name}...")
        new_column = insert_and_hide_column(ws, last_col, copied_data, last_row)
        
        new_column_name = xw.utils.col_name(new_column)
        print(f"Successfully inserted and hidden column: {new_column_name}")

        # Save and close
        wb.save()
        wb.close()
        print(f"✅ Modified file saved: {file_path}")
        
    except Exception as e:
        print(f"❌ Error processing file: {e}")
        if 'wb' in locals():
            wb.close()
        return

if __name__ == "__main__":
    main()
