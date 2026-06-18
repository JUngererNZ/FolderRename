import os
import sys
import json
import shutil
import pandas as pd
from datetime import datetime

# --- CONFIGURATION ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Path directly to the live shared Excel workbook for Forecast Tracking
CURRENT_DATA_PATH = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO TO ARRIVE AT DBN PORT\Z FORECAST\UPDATED FML FORECAST - 01.06.2026.xlsx"

# Unique snapshot tracking file for this forecast script
SNAPSHOT_PATH = os.path.join(SCRIPT_DIR, "forecast_snapshot.json")


def load_and_clean_forecast_data(file_path):
    """
    Reads the 'machine tracking' tab of the Excel file safely by creating a temporary copy
    to bypass aggressive OneDrive/Excel file locking environments.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Source file not found: {file_path}")

    # Path for a temporary copy in your project directory
    temp_file_path = os.path.join(SCRIPT_DIR, "~temp_forecast_updates.xlsx")
    
    try:
        # Create a local shadow copy of the workbook to dodge the lock
        shutil.copy2(file_path, temp_file_path)
        
        # Read from the temporary unlocked file copy (using 'machine tracking' tab)
        df_raw = pd.read_excel(temp_file_path, sheet_name="machine tracking", header=None, engine="openpyxl")
    finally:
        # Always remove the temporary file immediately after reading
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
    
    # Identify headers from row index 0 (Excel Row 1)
    headers = [str(x).strip() if pd.notna(x) else f"Unnamed: {i}" for i, x in enumerate(df_raw.iloc[0])]
    
    records = {}
    current_section = "General Forecast"
    
    # Iterate through data starting from row index 1 (Excel Row 2)
    for idx in range(1, len(df_raw)):
        row_cells = df_raw.iloc[idx].tolist()
        ba_cell = str(row_cells[0]).strip() if pd.notna(row_cells[0]) else ""
        
        # Skip empty rows
        if not ba_cell or ba_cell == "nan":
            continue
            
        # Detect Section Headers if any exist
        is_empty_rest = all(pd.isna(cell) or str(cell).strip() == "" for cell in row_cells[1:4])
        if is_empty_rest and not ba_cell.upper().startswith("BA"):
            current_section = ba_cell
            continue

        # Extract valid BA Number rows
        if ba_cell.upper().startswith("BA"):
            row_dict = {}
            for i, h in enumerate(headers):
                if i < len(row_cells):
                    val = row_cells[i]
                    
                    # Format dates cleanly if pandas parsed them as timestamps
                    if isinstance(val, (pd.Timestamp, datetime)):
                        row_dict[h] = val.strftime('%d-%m-%Y')
                    else:
                        row_dict[h] = str(val).strip() if pd.notna(val) else ""
            
            # Match exact Excel visual row layout (1-indexed row number)
            row_dict["_location"] = {
                "row_num": idx + 1,
                "section": current_section
            }
            records[ba_cell] = row_dict

    return records


def calculate_changes(old_snapshot, new_snapshot):
    """Compares baseline snapshot JSON with live Excel forecast workbook state."""
    added = []
    removed = []
    modified = {}

    # Check for additions and internal modifications
    for ba_num, new_data in new_snapshot.items():
        loc = new_data.get("_location", {"row_num": "Unknown", "section": "Unknown"})
        
        if ba_num not in old_snapshot:
            added.append((ba_num, new_data, loc))
        else:
            old_data = old_snapshot[ba_num]
            field_changes = {}
            
            for col, new_val in new_data.items():
                if col == "_location":
                    continue
                old_val = old_data.get(col, "")
                if str(old_val) != str(new_val):
                    field_changes[col] = {"from": old_val, "to": new_val}

            if field_changes:
                modified[ba_num] = {
                    "CUSTOMER": new_data.get("Customer", new_data.get("Client", "Unknown")),
                    "MODEL": new_data.get("Model", "Unknown"),
                    "LOCATION": loc,
                    "CHANGES": field_changes,
                }

    # Check for removals
    for ba_num, old_data in old_snapshot.items():
        if ba_num not in new_snapshot:
            loc = old_data.get("_location", {"row_num": "Unknown", "section": "Unknown"})
            removed.append((ba_num, old_data, loc))

    return added, removed, modified


def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scanning Live Forecast Workbook for updates...")

    try:
        current_state = load_and_clean_forecast_data(CURRENT_DATA_PATH)
    except Exception as e:
        print(f"Error processing current data file: {e}")
        sys.exit(1)

    if not os.path.exists(SNAPSHOT_PATH):
        print(f"-> Baseline snapshot profile not found. Creating local '{SNAPSHOT_PATH}'.")
        with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
            json.dump(current_state, f, indent=4)
        print("-> Snapshot saved successfully as initial baseline.")
        sys.exit(0)

    with open(SNAPSHOT_PATH, "r", encoding="utf-8") as f:
        historical_baseline = json.load(f)

    # Calculate difference profiles
    added, removed, modified = calculate_changes(historical_baseline, current_state)

    print("\n================== LIVE FORECAST CHANGE DETECTION REPORT ==================")
    
    if not added and not removed and not modified:
        print(" No changes detected since the last snapshot.")
    
    if added:
        print(f"\n[+] ADDED TO FORECAST ({len(added)}):")
        for ba_num, item, loc in added:
            print(f"  - BA: {ba_num} | Customer: {item.get('Customer', 'N/A')} | Model: {item.get('Model', 'N/A')}")
            print(f"    Excel Location -> Section: [{loc['section']}] | Row: {loc['row_num']}\n")

    if removed:
        print(f"\n[-] REMOVED FROM FORECAST ({len(removed)}):")
        for ba_num, item, loc in removed:
            print(f"  - BA: {ba_num} | Customer: {item.get('Customer', 'N/A')}")
            print(f"    Previous Location -> Section: [{loc['section']}] | Last Row: {loc['row_num']}\n")

    if modified:
        print(f"\n[*] MODIFIED CELLS ({len(modified)}):")
        for ba_num, details in modified.items():
            loc = details["LOCATION"]
            print(f"  - BA: {ba_num} (Customer: {details['CUSTOMER']} | Model: {details['MODEL']})")
            print(f"    Excel Location -> Section: [{loc['section']}] | Row: {loc['row_num']}")
            for field, values in details["CHANGES"].items():
                print(f"      * {field}: '{values['from']}' -> '{values['to']}'")
            print()
            
    print("==========================================================================\n")

    # Automatically save and advance the tracking baseline profile
    with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
        json.dump(current_state, f, indent=4)
    print("Local snapshot json base advanced to latest state.")
    
    sys.exit(0)


if __name__ == "__main__":
    main()