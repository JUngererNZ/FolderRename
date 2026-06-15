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
# CONFIGURATION
# -------------------------------
EXCEL_FILE = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO TO ARRIVE AT DBN PORT\VESSEL UPDATES.xlsx"
SHEET_NAME = "VESSEL ETA"
AISSTREAM_API_KEY = "4a90079dd212f4fc6ecf85c536477e0c974b8bb5"
MMSI_CONFIG_FILE = "vessel_mmsi_config.json"

# -------------------------------
# Load MMSI config
# -------------------------------
def load_mmsi_config():
    if os.path.exists(MMSI_CONFIG_FILE):
        with open(MMSI_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        print(f"⚠️ Config file {MMSI_CONFIG_FILE} not found. Using empty mapping.")
        return {}

VESSEL_MMSI_MAP = load_mmsi_config()

# -------------------------------
# Shadow copy for locked files
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
# Split vessel name
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
# Fetch live data (using API key as query param)
# -------------------------------
async def fetch_live_data(mmsi_list):
    if not mmsi_list:
        return {}
    
    uri = f"wss://stream.aisstream.io/v2?apiKey={AISSTREAM_API_KEY}"
    result = {}
    
    try:
        async with websockets.connect(uri, subprotocols=["graphql-ws"]) as websocket:
            await websocket.send(json.dumps({"type": "connection_init"}))
            init_resp = await websocket.recv()
            print(f"Connection response: {init_resp}")
            
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
                    if data.get("type") == "data" and "payload" in data:
                        vessel_data = data["payload"]["data"].get("vessels")
                        if vessel_data and vessel_data.get("mmsi"):
                            mmsi = vessel_data["mmsi"]
                            result[mmsi] = {
                                "timestamp": vessel_data.get("timestamp"),
                                "position_report": vessel_data.get("positionReport"),
                                "static_data": vessel_data.get("shipStaticData")
                            }
                except asyncio.TimeoutError:
                    continue
            
            print(f"Collected data for {len(result)} vessels.")
            return result
    except Exception as e:
        print(f"WebSocket error: {e}")
        return {}

# -------------------------------
# Extract vessel info from live data
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
# Determine status
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
# Main
# -------------------------------
async def main():
    print(f"Reading Excel file: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found at {EXCEL_FILE}")
        return
    
    shadow = get_readable_copy(EXCEL_FILE)
    if shadow:
        df = pd.read_excel(shadow, sheet_name=SHEET_NAME, header=1, dtype=str)
        os.unlink(shadow)
        print("Shadow copy removed.")
    else:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=1, dtype=str)
    
    df.columns = df.columns.str.strip()
    if "BA NUMBER" in df.columns:
        df = df.dropna(subset=["BA NUMBER"])
        df = df[~df["BA NUMBER"].astype(str).str.contains(r'JUNE|JULY|AUGUST|SEPTEMBER', case=False, na=False)]
    else:
        print("Column 'BA NUMBER' not found. Available:", df.columns.tolist())
        return
    
    required = ["BA NUMBER", "VESESL", "ETA DURBAN"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"Missing columns: {missing}")
        return
    
    print(f"\n✅ Using columns: BA NUMBER, VESESL, ETA DURBAN")
    
    df_filtered = df[required].dropna(subset=["VESESL", "ETA DURBAN"])
    df_filtered = df_filtered[~df_filtered["VESESL"].str.upper().str.contains("TBA", na=False)]
    
    if df_filtered.empty:
        print("No valid rows found.")
        return
    
    base_to_mmsi = {}
    pending_rows = []
    
    for _, row in df_filtered.iterrows():
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
    
    mmsi_list = [m for m in base_to_mmsi.values() if m and m > 0]
    live_data = {}
    if mmsi_list:
        print(f"Fetching live data for MMSIs: {mmsi_list}")
        live_data = await fetch_live_data(mmsi_list)
    else:
        print("No valid MMSI numbers; skipping live fetch.")
    
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
                "Status": row["reason"],
                "MMSI": None
            })
            continue
        
        base = row["base"]
        mmsi = base_to_mmsi.get(base)
        if mmsi and mmsi in live_data:
            live_info = extract_vessel_info(live_data, mmsi)
            status = compute_status(row["eta_file"], live_info)
            final_mmsi = mmsi
        else:
            live_info = {"current_port": "N/A", "destination": "N/A", "live_eta": "N/A"}
            status = "No MMSI or no live data"
            final_mmsi = mmsi if mmsi else None
        
        final.append({
            "BA": row["BA"],
            "Vessel (raw)": row["raw_vessel"],
            "Vessel Base": base,
            "Voyage Number": row["voyage"] if row["voyage"] else "",
            "ETA Durban (file)": row["eta_file"],
            "Current Port": live_info["current_port"],
            "Destination": live_info["destination"],
            "Live ETA": live_info["live_eta"],
            "Status": status,
            "MMSI": final_mmsi
        })
    
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