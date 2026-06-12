import requests
import json
import pandas as pd
import time
import re
import os
import shutil
import tempfile
from datetime import datetime

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
EXCEL_FILE = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO TO ARRIVE AT DBN PORT\VESSEL UPDATES.xlsx"
SHEET_NAME = "VESSEL ETA"
API_KEY = "0255ef2cb461087caad4c31fa4b1a762ff98f2d9a8babb7701d2a5ca5a2de6d1"   # Your VesselAPI key
BASE_URL = "https://api.vesselapi.com/v1"   # Correct base for VesselAPI
HEADERS = {"Authorization": f"Bearer {API_KEY}"}
MMSI_CONFIG_FILE = "vessel_mmsi_config.json"

# ------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------
def load_mmsi_config():
    if os.path.exists(MMSI_CONFIG_FILE):
        with open(MMSI_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        print(f"⚠️ Config file {MMSI_CONFIG_FILE} not found.")
        return {}

def get_readable_copy(source_path, max_retries=3):
    for attempt in range(max_retries):
        try:
            ext = os.path.splitext(source_path)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp_path = tmp.name
            shutil.copy2(source_path, tmp_path)
            print(f"Created shadow copy: {tmp_path}")
            return tmp_path
        except (PermissionError, OSError) as e:
            print(f"Attempt {attempt+1}/{max_retries} failed: {e}")
            time.sleep(1)
    return None

def split_vessel_name(raw_name):
    if pd.isna(raw_name) or raw_name == "":
        return None, None
    pattern = r'\s+(V\.|Voy\.)\s*(.*)$'
    match = re.search(pattern, raw_name, re.IGNORECASE)
    if match:
        voyage = f"{match.group(1)} {match.group(2)}".strip()
        base = raw_name[:match.start()].strip()
        return base, voyage
    return raw_name.strip(), None

def get_vessel_static(mmsi):
    """Fetch static vessel data (name, IMO, etc.)"""
    url = f"{BASE_URL}/vessel/{mmsi}?filter.idType=mmsi"
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            return r.json().get("vessel", {})
    except Exception as e:
        print(f"Static error for {mmsi}: {e}")
    return None

def get_vessel_position(mmsi):
    """Fetch live position (latitude, longitude)"""
    url = f"{BASE_URL}/position/{mmsi}?filter.idType=mmsi"
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        print(f"Position error for {mmsi}: {e}")
    return None

def get_vessel_voyage(mmsi):
    """Fetch destination and ETA"""
    url = f"{BASE_URL}/voyage/{mmsi}?filter.idType=mmsi"
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        print(f"Voyage error for {mmsi}: {e}")
    return None

def get_vessel_info(mmsi):
    """Combine all data"""
    static = get_vessel_static(mmsi)
    position = get_vessel_position(mmsi)
    voyage = get_vessel_voyage(mmsi)

    lat = None
    lon = None
    if position and "position" in position:
        lat = position["position"].get("latitude")
        lon = position["position"].get("longitude")
    destination = None
    eta = None
    if voyage and "voyage" in voyage:
        destination = voyage["voyage"].get("destination")
        eta = voyage["voyage"].get("eta")

    return {
        "name": static.get("name") if static else None,
        "latitude": lat,
        "longitude": lon,
        "destination": destination,
        "eta": eta,
        "static_ok": static is not None,
        "position_ok": position is not None,
        "voyage_ok": voyage is not None,
    }

# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
def main():
    print(f"Reading Excel: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        print("File not found.")
        return

    shadow = get_readable_copy(EXCEL_FILE)
    if shadow:
        df = pd.read_excel(shadow, sheet_name=SHEET_NAME, header=1, dtype=str)
        os.unlink(shadow)
    else:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=1, dtype=str)

    df.columns = df.columns.str.strip()
    if "BA NUMBER" not in df.columns:
        print("Column 'BA NUMBER' missing.")
        return

    df = df.dropna(subset=["BA NUMBER", "VESESL", "ETA DURBAN"])
    df = df[~df["VESESL"].str.upper().str.contains("TBA|JUNE|JULY|AUGUST|SEPTEMBER", na=False)]

    if df.empty:
        print("No valid rows.")
        return

    mmsi_map = load_mmsi_config()
    results = []

    for _, row in df.iterrows():
        ba = row["BA NUMBER"]
        raw_vessel = row["VESESL"]
        eta_file = row["ETA DURBAN"]
        base, voyage = split_vessel_name(raw_vessel)

        mmsi = mmsi_map.get(base)
        if not mmsi:
            results.append({
                "BA": ba,
                "Vessel (raw)": raw_vessel,
                "Vessel Base": base,
                "Voyage Number": voyage or "",
                "ETA Durban (file)": eta_file,
                "MMSI": None,
                "Name": None,
                "Latitude": None,
                "Longitude": None,
                "Destination": None,
                "Live ETA": None,
                "Status": "No MMSI in config"
            })
            continue

        print(f"Fetching data for {base} (MMSI {mmsi})...")
        info = get_vessel_info(mmsi)

        # Build status message
        if info["static_ok"]:
            status = "Static OK"
            if info["position_ok"]:
                status += ", position OK"
            else:
                status += ", no position (AIS off?)"
            if info["voyage_ok"]:
                status += ", voyage OK"
            else:
                status += ", no voyage data"
        else:
            status = "Static data not found (invalid MMSI?)"

        results.append({
            "BA": ba,
            "Vessel (raw)": raw_vessel,
            "Vessel Base": base,
            "Voyage Number": voyage or "",
            "ETA Durban (file)": eta_file,
            "MMSI": mmsi,
            "Name": info["name"],
            "Latitude": info["latitude"],
            "Longitude": info["longitude"],
            "Destination": info["destination"],
            "Live ETA": info["eta"],
            "Status": status
        })
        time.sleep(0.5)  # be polite

    # Save reports
    out_df = pd.DataFrame(results)
    out_df.to_csv("vessel_report.csv", index=False)
    with open("vessel_report.json", "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, default=str)

    print("\n✅ Reports saved: vessel_report.csv / vessel_report.json")
    print("\nSummary:")
    print(out_df.to_string(index=False))

if __name__ == "__main__":
    main()