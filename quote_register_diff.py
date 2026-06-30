import os
import sys
import json
import shutil
import pandas as pd
from datetime import datetime

# --- CONFIGURATION 18062026---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Live shared Excel workbook path for Quote Register
CURRENT_DATA_PATH = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\FML QUOTE REGISTER\QUOTE REGISTER AS OF 24-06-2026.xlsx"

# Local snapshot tracking file adjusted for this script
SNAPSHOT_PATH = os.path.join(SCRIPT_DIR, "quote_snapshot.json")


def load_and_clean_quote_data(file_path):
    """
    Reads the data tab of the Quote Register safely by creating a temporary copy
    to bypass aggressive OneDrive/Excel file locking environments.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Source file not found: {file_path}")

    # Path for a temporary copy in your project directory
    temp_file_path = os.path.join(SCRIPT_DIR, "~temp_quote_register.xlsx")
    
    try:
        # Create a local shadow copy of the workbook to dodge the lock
        shutil.copy2(file_path, temp_file_path)
        
        # NOTE: Using sheet_name=0 (first sheet) as a fallback if the exact name changes,
        # but you can change this to a string like "QUOTES" or "REGISTER" if explicitly named.
        df_raw = pd.read_excel(temp_file_path, sheet_name=0, header=None, engine="openpyxl")
    finally:
        # Always remove the temporary file immediately after reading
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
    
    # Identify headers from row index 1 (Excel Row 2) -> Adjust index if headers are on Row 1 (index 0)
    headers = [str(x).strip() if pd.notna(x) else f"Unnamed: {i}" for i, x in enumerate(df_raw.iloc[1])]
    
    records = {}
    current_section = "General"
    
    # Iterate through data starting from row index 2 (Excel Row 3)
    for idx in range(2, len(df_raw)):
        row_cells = df_raw.iloc[idx].tolist()
        key_cell = str(row_cells[0]).strip() if pd.notna(row_cells[0]) else ""
        
        # Skip empty rows completely
        if not key_cell or key_cell == "nan":
            continue
            
        # Detect Section Headers if applicable (e.g., monthly breakdowns)
        # If your register doesn't use section headers, this won't interfere as long as quotes don't match criteria
        is_empty_rest = all(pd.isna(cell) or str(cell).strip() == "" for cell in row_cells[1:4])
        if is_empty_rest and not (key_cell.upper().startswith("Q") or key_cell.isdigit()):
            current_section = key_cell
            continue

        # Treat row as a valid line item record
        row_dict = {}
        for i, h in enumerate(headers):
            val = row_cells[i]
            
            if isinstance(val, (pd.Timestamp, datetime)):
                row_dict[h] = val.strftime('%d-%m-%Y')
            else:
                row_dict[h] = str(val).strip() if pd.notna(val) else ""
        
        # Store layout tracking data
        row_dict["_location"] = {
            "row_num": idx + 1,
            "section": current_section
        }
        records[key_cell] = row_dict

    return records


def calculate_changes(old_snapshot, new_snapshot):
    """Compares baseline snapshot JSON with live Excel workbook state."""
    added = []
    removed = []
    modified = {}

    # Check for additions and modifications
    for key_id, new_data in new_snapshot.items():
        loc = new_data.get("_location", {"row_num": "Unknown", "section": "Unknown"})
        
        if key_id not in old_snapshot:
            added.append((key_id, new_data, loc))
        else:
            old_data = old_snapshot[key_id]
            field_changes = {}
            
            for col, new_val in new_data.items():
                if col == "_location":
                    continue
                old_val = old_data.get(col, "")
                if str(old_val) != str(new_val):
                    field_changes[col] = {"from": old_val, "to": new_val}

            if field_changes:
                # Uses fallback generic field names if client/description columns aren't standard
                modified[key_id] = {
                    "CLIENT": new_data.get("CLIENT", new_data.get("CUSTOMER", "N/A")),
                    "LOCATION": loc,
                    "CHANGES": field_changes,
                }

    # Check for removals
    for key_id, old_data in old_snapshot.items():
        if key_id not in new_snapshot:
            loc = old_data.get("_location", {"row_num": "Unknown", "section": "Unknown"})
            removed.append((key_id, old_data, loc))

    return added, removed, modified


def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scanning Live Quote Register for updates...")

    try:
        current_state = load_and_clean_quote_data(CURRENT_DATA_PATH)
    except Exception as e:
        print(f"ERROR processing current data file: {e}")
        # Explicitly exit with failure status (Exit Code 1)
        sys.exit(1)

    if not os.path.exists(SNAPSHOT_PATH):
        print(f"-> Baseline snapshot profile not found. Creating local '{SNAPSHOT_PATH}'.")
        with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
            json.dump(current_state, f, indent=4)
        print("-> Snapshot saved successfully as initial baseline.")
        # Exit with success status (Exit Code 0)
        sys.exit(0)

    with open(SNAPSHOT_PATH, "r", encoding="utf-8") as f:
        historical_baseline = json.load(f)

    # Calculate difference profiles
    added, removed, modified = calculate_changes(historical_baseline, current_state)

    print("\n================== LIVE QUOTE REGISTER CHANGE DETECTION REPORT ==================")
    
    if not added and not removed and not modified:
        print(" No changes detected since the last snapshot.")
    
    if added:
        print(f"\n[+] ADDED TO REGISTER ({len(added)}):")
        for key_id, item, loc in added:
            client = item.get('CLIENT', item.get('CUSTOMER', 'Unknown'))
            print(f"  - Key ID/Quote: {key_id} | Client: {client}")
            print(f"    Excel Location -> Section: [{loc['section']}] | Row: {loc['row_num']}\n")

    if removed:
        print(f"\n[-] REMOVED FROM REGISTER ({len(removed)}):")
        for key_id, item, loc in removed:
            client = item.get('CLIENT', item.get('CUSTOMER', 'Unknown'))
            print(f"  - Key ID/Quote: {key_id} | Client: {client}")
            print(f"    Previous Location -> Section: [{loc['section']}] | Last Row: {loc['row_num']}\n")

    if modified:
        print(f"\n[*] MODIFIED CELLS ({len(modified)}):")
        for key_id, details in modified.items():
            loc = details["LOCATION"]
            print(f"  - Key ID/Quote: {key_id} ({details['CLIENT']})")
            print(f"    Excel Location -> Section: [{loc['section']}] | Row: {loc['row_num']}")
            for field, values in details["CHANGES"].items():
                print(f"      * {field}: '{values['from']}' -> '{values['to']}'")
            print()
            
    print("=================================================================================\n")

    # Automatically save and advance the tracking baseline profile
    try:
        with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
            json.dump(current_state, f, indent=4)
        print("Local snapshot json base advanced to latest state.")
    except Exception as e:
        print(f"ERROR saving snapshot file: {e}")
        sys.exit(1)
    
    # Exit cleanly with success status (Exit Code 0)
    sys.exit(0)


if __name__ == "__main__":
    main()