import requests
import json

# --------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------
API_KEY = "0255ef2cb461087caad4c31fa4b1a762ff98f2d9a8babb7701d2a5ca5a2de6d1"  # Your API key
BASE_URL = "https://vesselapi.com"                                            # Correct base URL
HEADERS = {"Authorization": f"Bearer {API_KEY}"}

def get_vessel_static_info(mmsi):
    """Retrieve static vessel information (name, IMO, dimensions, etc.)"""
    url = f"{BASE_URL}/vessel/{mmsi}"
    params = {"filter.idType": "mmsi"}
    try:
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        return response.json().get("vessel")
    except Exception as e:
        print(f"Error fetching static data for MMSI {mmsi}: {e}")
        return None

def get_vessel_position(mmsi):
    """Retrieve the most recent AIS position for a vessel"""
    url = f"{BASE_URL}/vessel/{mmsi}/position"
    params = {"filter.idType": "mmsi"}
    try:
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching position for MMSI {mmsi}: {e}")
        return None

def get_vessel_eta(mmsi):
    """Retrieve crew-reported destination and ETA for a vessel"""
    url = f"{BASE_URL}/vessel/{mmsi}/eta"
    params = {"filter.idType": "mmsi"}
    try:
        response = requests.get(url, headers=HEADERS, params=params)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching ETA for MMSI {mmsi}: {e}")
        return None

def get_vessel_info(mmsi):
    """Combine static, position, and ETA data for a complete picture."""
    print(f"\n🚢 Fetching data for MMSI: {mmsi}")
    
    # Get static info (works even if vessel is not transmitting)
    static = get_vessel_static_info(mmsi)
    if static:
        print(f"✅ Vessel Name: {static.get('name')}, IMO: {static.get('imo')}")
    
    # Get live position (requires recent AIS transmission)
    position = get_vessel_position(mmsi)
    if position:
        print(f"📍 Last Position: {position.get('latitude')}, {position.get('longitude')} at {position.get('timestamp')}")
    else:
        print("⚠️ No live position – vessel may be out of range or AIS off.")
    
    # Get crew-reported ETA
    eta = get_vessel_eta(mmsi)
    if eta:
        print(f"🗺️ Destination: {eta.get('destination')}, ETA: {eta.get('eta')}, Draught: {eta.get('draught')}m")
    else:
        print("⚠️ No ETA reported.")
    
    return {"static": static, "position": position, "eta": eta}

# --------------------------------------------------------
# YOUR MMSI LIST (FROM vessel_mmsi_config.json)
# --------------------------------------------------------
mmsi_list = [636022934, 440114000, 258628000, 441390000]

# Fetch data for all vessels
all_vessel_data = {}
for mmsi in mmsi_list:
    all_vessel_data[mmsi] = get_vessel_info(mmsi)

# Save to JSON for later processing
with open("vessel_data.json", "w") as f:
    json.dump(all_vessel_data, f, indent=2, default=str)

print("\n✅ Data saved to vessel_data.json")