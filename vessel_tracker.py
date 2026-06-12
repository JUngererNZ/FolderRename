import asyncio
import json
import pandas as pd
import time
import re
import os
import shutil
import tempfile
from datetime import datetime
import websockets

# -------------------------------
# CONFIGURATION 4
# -------------------------------
EXCEL_FILE = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO TO ARRIVE AT DBN PORT\VESSEL UPDATES.xlsx"
SHEET_NAME = "VESSEL ETA"

AISSTREAM_API_KEY = "4a90079dd212f4fc6ecf85c536477e0c974b8bb5"

# MAPPING: vessel base name -> MMSI (YOU MUST FILL THIS!)
VESSEL_MMSI_MAP = {
    "ASIAN EMPIRE": 440114000,
    "HAN JIANG KOU": 0,
    "MORNING CELLO": 0,
    "HOEGH TRACER": 0,
}

# -------------------------------
# Helper: create a shadow copy
# -------------------------------
def get_readable_copy(source_path, max_retries=3, retry_delay=1):
    for attempt in range(max_retries):
        try:
            ext = os.path.splitext(source_path)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_file:
                tmp_path = tmp_file.name
            shutil.copy2(source_path, tmp_path)
            print(f"Created shadow copy: {tmp_path}")
            return tmp_path
        except (PermissionError, OSError) as e:
            print(f"Attempt {attempt+1}/{max_retries} failed: {e}")
            time.sleep(retry_delay)
    print("Failed to create shadow copy after multiple attempts.")
    return None

# -------------------------------
# Helper: split vessel name
# -------------------------------
def split_vessel_name(raw_name):
    if pd.isna(raw_name) or raw_name == "":
        return None, None
    pattern = r'\s+(V\.|Voy\.)\s*(.*)$'
    match = re.search(pattern, raw_name, re.IGNORECASE)
    if match:
        voyage = f"{match.group(1)} {match.group(2)}".strip()
        base = raw_name[:match.start()].strip()
        return base, voyage
    else:
        return raw_name.strip(), None

# -------------------------------
# Fetch live data via AISStream
# -------------------------------
async def fetch_live_data(mmsi_list):
    if not mmsi_list:
        return {}
    
    uri = "wss://stream.aisstream.io/v2/stream"
    result = {}
    
    async with websockets.connect(uri, subprotocols=["graphql-ws"]) as websocket:
        await websocket.send(json.dumps({
            "type": "connection_init",
            "payload": {"headers": {"X-API-Key": AISSTREAM_API_KEY}}
        }))
        init_resp = await websocket.recv()
        print(f"Connection init response: {init_resp}")
        
        subscribe_msg = {
            "type": "subscribe",
            "id": "1",
            "payload": {
                "query": f"""subscription {{
                    vessels(
                        mmsi: {json.dumps(mmsi_list)}
                        messageTypes: ["PositionReport", "ShipStaticData"]
                    ) {{
                        mmsi
                        timestamp
                        positionReport {{
                            latitude
                            longitude
                            sog
                            cog
                        }}
                        shipStaticData {{
                            destination
                            eta
                        }}
                    }}
                }}"""
            }
        }
        await websocket.send(json.dumps(subscribe_msg))
        print("Subscribed, waiting for data...")
        
        start_time = time.time()
        timeout = 15
        
        while time.time() - start_time < timeout:
            try:
                message = await asyncio.wait_for(websocket.recv(), timeout=1)
                data = json.loads(message)
                if 'payload' in data and 'data' in data['payload']:
                    vessel = data['payload']['data']['vessels']
                    if vessel and vessel.get('mmsi'):
                        mmsi = vessel['mmsi']
                        result[mmsi] = {
                            "timestamp": vessel.get('timestamp'),
                            "position_report": vessel.get('positionReport'),
                            "static_data": vessel.get('shipStaticData')
                        }
            except asyncio.TimeoutError:
                continue
        
        print(f"Collected data for {len(result)} vessels.")
        return result

# -------------------------------
# Convert AISStream data
# -------------------------------
def extract_vessel_info(live_data, mmsi):
    if mmsi not in live_data:
        return {"current_port": "N/A", "destination": "N/A", "live_eta": "N/A"}
    
    data = live_data[mmsi]
    static = data.get("static_data", {}) or {}
    pos = data.get("position_report", {}) or {}
    
    destination = static.get("destination", "N/A")
    eta = static.get("eta", "N/A")
    
    lat = pos.get("latitude")
    lon = pos.get("longitude")
    if lat and lon:
        current_port = f"{lat:.3f}, {lon:.3f}"
    else:
        current_port = "N/A"
    
    return {
        "current_port": current_port,
        "destination": destination,
        "live_eta": eta,
        "sog": pos.get("sog"),
        "cog": pos.get("cog")
    }

# -------------------------------
# Compare file ETA with live info
# -------------------------------
def compute_status(file_eta_str, live_info):
    try:
        eta_file = datetime.strptime(file_eta_str.strip(), "%d-%m-%Y").date()
    except:
        return f"Invalid ETA format: {file_eta_str}"
    
    dest = live_info.get("destination", "").lower()
    if "durban" in dest:
        return f"En route to Durban (live ETA: {live_info['live_eta']})"
    elif live_info["current_port"] != "N/A":
        return f"Last known position: {live_info['current_port']}"
    else:
        return "No live position available"

# -------------------------------
# Main routine
# -------------------------------
async def main():
    print(f"Reading Excel file: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found at {EXCEL_FILE}")
        return
    
    # Shadow copy
    shadow = get_readable_copy(EXCEL_FILE)
    if shadow:
        # Read with header=1 (second row) to get correct column names
        df = pd.read_excel(shadow, sheet_name=SHEET_NAME, header=1, dtype=str)
        os.unlink(shadow)
        print("Shadow copy removed.")
    else:
        print("Unable to create shadow copy. Attempting to read original directly...")
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=1, dtype=str)
    
    # Remove rows where BA NUMBER is empty or contains month/year like "JUNE 2026"
    # The BA column after header is likely named "BA NUMBER" (as shown in row 1)
    # We'll rename columns by stripping spaces and standardizing
    df.columns = df.columns.str.strip()
    
    # Filter out rows where BA NUMBER is NaN or is a month string (like JUNE 2026)
    if "BA NUMBER" in df.columns:
        df = df.dropna(subset=["BA NUMBER"])
        df = df[~df["BA NUMBER"].astype(str).str.contains(r'JUNE|JULY|AUGUST|SEPTEMBER', case=False, na=False)]
    else:
        print("Column 'BA NUMBER' not found after reading with header=1. Available columns:", df.columns.tolist())
        return
    
    # Now we have the correct data rows. Also, some rows may have vessel names like "TBA" - we'll keep them for now
    # but later skip if vessel name is missing or TBA.
    
    # Check we have the required columns
    required = ["BA NUMBER", "VESESL", "ETA DURBAN"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"Missing required columns: {missing}")
        print("Available columns:", df.columns.tolist())
        return
    
    print(f"\n✅ Using columns: BA NUMBER, VESESL, ETA DURBAN")
    
    # Filter rows that have vessel name and ETA (drop rows where vessel is empty or TBA)
    df_filtered = df[required].dropna(subset=["VESESL", "ETA DURBAN"])
    df_filtered = df_filtered[~df_filtered["VESESL"].str.upper().str.contains("TBA", na=False)]
    
    if df_filtered.empty:
        print("No valid rows found after filtering.")
        return
    
    # Build list of unique vessel base names
    base_to_mmsi = {}
    pending_rows = []
    
    for idx, row in df_filtered.iterrows():
        ba = row["BA NUMBER"]
        raw_vessel = row["VESESL"]
        eta_file = row["ETA DURBAN"]
        base, voyage = split_vessel_name(raw_vessel)
        
        if not base:
            pending_rows.append({
                "BA": ba,
                "raw_vessel": raw_vessel,
                "base": None,
                "voyage": None,
                "eta_file": eta_file,
                "skip": True,
                "reason": "Could not extract vessel name"
            })
            continue
        
        if base not in base_to_mmsi:
            mmsi = VESSEL_MMSI_MAP.get(base)
            if not mmsi:
                print(f"⚠️ Warning: No MMSI found for '{base}'. Live data will be skipped.")
            base_to_mmsi[base] = mmsi
        
        pending_rows.append({
            "BA": ba,
            "raw_vessel": raw_vessel,
            "base": base,
            "voyage": voyage,
            "eta_file": eta_file,
            "skip": False
        })
    
    # Collect live data for MMSIs that exist (>0)
    mmsi_list = [m for m in base_to_mmsi.values() if m and m > 0]
    live_data = {}
    if mmsi_list:
        print(f"Fetching live data for MMSIs: {mmsi_list}")
        live_data = await fetch_live_data(mmsi_list)
    else:
        print("No valid MMSI numbers provided; skipping live fetch.")
    
    # Build final results
    final = []
    for row in pending_rows:
        if row["skip"]:
            final.append({
                "BA": row["BA"],
                "Vessel (raw)": row["raw_vessel"],
                "Vessel Base": None,
                "Voyage Number": "",
                "ETA Durban (file)": row["eta_file"],
                "Current Port": "N/A",
                "Destination": "N/A",
                "Live ETA": "N/A",
                "Status": row["reason"]
            })
            continue
        
        base = row["base"]
        mmsi = base_to_mmsi.get(base)
        if mmsi and mmsi in live_data:
            live_info = extract_vessel_info(live_data, mmsi)
            status = compute_status(row["eta_file"], live_info)
        else:
            live_info = {"current_port": "N/A", "destination": "N/A", "live_eta": "N/A"}
            status = "No MMSI or no live data (vessel not transmitting?)"
        
        final.append({
            "BA": row["BA"],
            "Vessel (raw)": row["raw_vessel"],
            "Vessel Base": base,
            "Voyage Number": row["voyage"] if row["voyage"] else "",
            "ETA Durban (file)": row["eta_file"],
            "Current Port": live_info["current_port"],
            "Destination": live_info["destination"],
            "Live ETA": live_info["live_eta"],
            "Status": status
        })
    
    # Save reports
    output_json = "vessel_report.json"
    output_csv = "vessel_report.csv"
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(final, f, indent=2, default=str)
    pd.DataFrame(final).to_csv(output_csv, index=False)
    
    print(f"\n✅ Reports saved: {output_json} and {output_csv}")
    print("\nSummary:")
    print(pd.DataFrame(final).to_string(index=False))

if __name__ == "__main__":
    asyncio.run(main())