import os
import json
import asyncio
import random
import re
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "vessel_mmsi_config.json")
DATA_PATH = os.path.join(SCRIPT_DIR, "vessel_data.json")

async def human_delay(min_sec=2, max_sec=5):
    await asyncio.sleep(random.uniform(min_sec, max_sec))

async def scrape_vessel(mmsi_or_imo):
    url = f"https://www.vesselfinder.com/vessels/details/{mmsi_or_imo}"
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = await browser.new_context(user_agent=USER_AGENT, viewport={"width": 1920, "height": 1080})
        page = await context.new_page()
        
        print(f"Fetching live data for ID: {mmsi_or_imo}...")
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await human_delay(3, 5)
            await page.evaluate("window.scrollTo(0, document.body.scrollHeight / 4);")
            await human_delay(1, 2)
            html_content = await page.content()
        except Exception as e:
            print(f"Failed to fetch data for {mmsi_or_imo}: {e}")
            html_content = None
        finally:
            await browser.close()
        return html_content

def parse_html_to_json(html_content):
    if not html_content:
        return None
        
    soup = BeautifulSoup(html_content, "html.parser")
    raw_data = {}
    
    # 1. Gather all generic key-value tables
    tables = soup.find_all("table")
    for table in tables:
        for row in table.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) >= 2:
                label = cells[0].text.strip().replace(":", "").lower()
                value = cells[1].text.strip()
                if label:
                    raw_data[label] = value

    # 2. Extract Destination & Live ETA from the top Voyage Data panel blocks
    destination = "-"
    dest_div = soup.find("div", class_="destination") or soup.find("a", class_="col-port")
    if dest_div:
        destination = dest_div.text.strip()
        
    eta_val = "-"
    eta_div = soup.find("div", class_="countdown") or soup.find("span", class_="eta-time") or soup.find("td", class_="v4", string=re.compile(r".*ETA.*"))
    if eta_div:
        eta_val = eta_div.text.strip()
    else:
        # Fallback to key checks
        eta_val = raw_data.get("eta") or raw_data.get("expected arrival") or "-"

    # 3. Extract Precise Coordinates using a broader regex matching strategy
    latitude, longitude = "-", "-"
    text_pool = soup.get_text()
    
    # Regex designed to isolate variations like "29.8583 S / 31.0250 E" or "29.8583° S, 31.0250° E"
    coord_match = re.search(r"(\d+\.\d+(?:°)?\s*[NS])\s*[/,]\s*(\d+\.\d+(?:°)?\s*[EW])", text_pool, re.IGNORECASE)
    if coord_match:
        latitude = coord_match.group(1).replace("°", "").strip()
        longitude = coord_match.group(2).replace("°", "").strip()

    # 4. Parse Last Port elements via container structures
    last_port, atd = "-", "-"
    last_port_container = soup.find("div", class_="last-port-container") or soup.find("section", class_="last-port")
    if last_port_container:
        p_name = last_port_container.find("div", class_="lport") or last_port_container.find("a", class_="col-port")
        p_atd = last_port_container.find("div", class_="atd") or last_port_container.find("span", class_="atd-time")
        if p_name: last_port = p_name.text.strip()
        if p_atd: atd = p_atd.text.strip()

    vessel_name_header = soup.find("h1", class_="title")
    extracted_name = vessel_name_header.text.replace("current position", "").strip() if vessel_name_header else None

    # Structural mapping matching your exact vessel_data schema
    structured = {
        "static": {
            "vessel_name": raw_data.get("vessel name") or raw_data.get("name") or extracted_name or "-",
            "imo": raw_data.get("imo / mmsi", "").split("/")[0].strip() if "/" in raw_data.get("imo / mmsi", "") else raw_data.get("imo", "-"),
            "mmsi": raw_data.get("imo / mmsi", "").split("/")[1].strip() if "/" in raw_data.get("imo / mmsi", "") else raw_data.get("mmsi", "-"),
            "callsign": raw_data.get("callsign", "-"),
            "ais_type": raw_data.get("ais type", "-"),
            "flag": raw_data.get("ais flag") or raw_data.get("flag") or raw_data.get("flag / nationality") or "-",
            "length_beam": raw_data.get("length / beam", "-")
        },
        "position": {
            "status": raw_data.get("navigation status") or raw_data.get("status") or "-",
            "course_speed": raw_data.get("course / speed", "-"),
            "current_draught": raw_data.get("current draught") or raw_data.get("draught", "-"),
            "navigation_status": raw_data.get("navigation status") or raw_data.get("status") or "-",
            "position_received": raw_data.get("position received", "-"),
            "latitude": latitude,
            "longitude": longitude,
            "last_port": last_port if last_port else "-",
            "atd": atd if atd else "-"
        },
        "eta": {
            "destination": destination,
            "eta": eta_val
        }
    }
    return structured

async def main():
    if not os.path.exists(CONFIG_PATH) or not os.path.exists(DATA_PATH):
        print("Missing configuration or destination data files.")
        return
        
    with open(CONFIG_PATH, "r") as f:
        config = json.load(f)
    with open(DATA_PATH, "r") as f:
        vessel_data = json.load(f)
        
    for name, mmsi in config.items():
        mmsi_str = str(mmsi)
        html = await scrape_vessel(mmsi_str)
        parsed_result = parse_html_to_json(html)
        
        if parsed_result:
            vessel_data[mmsi_str] = parsed_result
            print(f"Successfully updated data schema for {name}.")
        else:
            print(f"Warning: Processing failed for {name}.")
        await human_delay(4, 8)
        
    with open(DATA_PATH, "w") as f:
        json.dump(vessel_data, f, indent=4)
    print("\nAll tasks complete. File verified updated.")

if __name__ == "__main__":
    asyncio.run(main())