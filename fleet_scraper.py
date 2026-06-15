import os
import json
import asyncio
import random
from playwright.async_api import async_playwright

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "vessel_mmsi_config.json")
DATA_PATH = os.path.join(SCRIPT_DIR, "vessel_data.json")

# ------------------------------------------------------------
# Helper: human‑like delay
# ------------------------------------------------------------
async def human_delay(min_sec=2, max_sec=5):
    await asyncio.sleep(random.uniform(min_sec, max_sec))

# ------------------------------------------------------------
# Scrape one vessel using JavaScript inside the page
# ------------------------------------------------------------
async def scrape_vessel_with_js(mmsi_or_imo):
    """
    Fetch live data for a vessel using its MMSI or IMO.
    Returns a structured dict with static, position, and eta info.
    """
    url = f"https://www.vesselfinder.com/vessels/details/{mmsi_or_imo}"
    async with async_playwright() as p:
        # Launch browser (headless = True for no GUI, set False for debugging)
        browser = await p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = await browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()
        
        print(f"Fetching live data for ID: {mmsi_or_imo}...")
        try:
            # Go to the vessel page and wait for content to load
            await page.goto(url, wait_until="networkidle", timeout=30000)
            await page.wait_for_selector("table", timeout=5000)
            await human_delay(2, 4)  # extra safety
            
            # Extract data using JavaScript (runs inside the browser)
            data = await page.evaluate("""() => {
                // Helper to get text from a selector
                function getText(selector) {
                    let el = document.querySelector(selector);
                    return el ? el.innerText.trim() : null;
                }
                
                // ----- Latitude / Longitude -----
                let lat = null, lon = null;
                // Look for common coordinate containers
                let coordSpan = document.querySelector(".lat-lon, .coords, [class*='coord']");
                if (coordSpan) {
                    let txt = coordSpan.innerText;
                    let match = txt.match(/(\\d+\\.\\d+)\\s*([NS]),?\\s*(\\d+\\.\\d+)\\s*([EW])/i);
                    if (match) {
                        lat = parseFloat(match[1]) * (match[2].toUpperCase() === 'S' ? -1 : 1);
                        lon = parseFloat(match[3]) * (match[4].toUpperCase() === 'W' ? -1 : 1);
                    }
                }
                // Fallback: look for any pattern like "29.8583 S / 31.0250 E" in text
                if (lat === null) {
                    let bodyText = document.body.innerText;
                    let fallbackMatch = bodyText.match(/(\\d+\\.\\d+)\\s*([NS])\\s*[\\/\\,]?\\s*(\\d+\\.\\d+)\\s*([EW])/i);
                    if (fallbackMatch) {
                        lat = parseFloat(fallbackMatch[1]) * (fallbackMatch[2].toUpperCase() === 'S' ? -1 : 1);
                        lon = parseFloat(fallbackMatch[3]) * (fallbackMatch[4].toUpperCase() === 'W' ? -1 : 1);
                    }
                }
                
                // ----- Destination & ETA -----
                let dest = getText(".destination a, .col-port, .dest");
                let etaElem = document.querySelector(".countdown, .eta-time");
                if (!etaElem) {
                    // Try table cell after "ETA" label
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toUpperCase() === "ETA" && cell.nextElementSibling) {
                            etaElem = cell.nextElementSibling;
                            break;
                        }
                    }
                }
                let eta = etaElem ? etaElem.innerText.trim() : null;
                
                // ----- Speed, draught, position age, last port -----
                let speed = null;
                let speedElem = document.querySelector(".speed");
                if (!speedElem) {
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toLowerCase() === "speed" && cell.nextElementSibling) {
                            speedElem = cell.nextElementSibling;
                            break;
                        }
                    }
                }
                if (speedElem) {
                    let speedText = speedElem.innerText;
                    let match = speedText.match(/(\\d+\\.?\\d*)/);
                    if (match) speed = match[1];
                }
                
                let draught = null;
                let draughtElem = document.querySelector(".draught");
                if (!draughtElem) {
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toLowerCase() === "current draught" && cell.nextElementSibling) {
                            draughtElem = cell.nextElementSibling;
                            break;
                        }
                    }
                }
                if (draughtElem) {
                    let draughtText = draughtElem.innerText;
                    let match = draughtText.match(/(\\d+\\.?\\d*)/);
                    if (match) draught = match[1];
                }
                
                let posAge = getText(".position-received");
                if (!posAge) {
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toLowerCase() === "position received" && cell.nextElementSibling) {
                            posAge = cell.nextElementSibling.innerText.trim();
                            break;
                        }
                    }
                }
                
                let lastPort = getText(".last-port a, .lport");
                if (!lastPort) {
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toLowerCase() === "last port" && cell.nextElementSibling) {
                            lastPort = cell.nextElementSibling.innerText.trim();
                            break;
                        }
                    }
                }
                
                let atd = null;
                let atdElem = document.querySelector(".atd");
                if (!atdElem) {
                    let cells = Array.from(document.querySelectorAll("td, th"));
                    for (let cell of cells) {
                        if (cell.innerText.trim().toUpperCase() === "ATD" && cell.nextElementSibling) {
                            atdElem = cell.nextElementSibling;
                            break;
                        }
                    }
                }
                if (atdElem) atd = atdElem.innerText.trim();
                
                // ----- Static vessel data from tables -----
                let staticData = {};
                let tables = document.querySelectorAll("table");
                for (let table of tables) {
                    let rows = table.querySelectorAll("tr");
                    for (let row of rows) {
                        let cells = row.querySelectorAll("td");
                        if (cells.length >= 2) {
                            let key = cells[0].innerText.trim().toLowerCase().replace(/:/g, '');
                            let val = cells[1].innerText.trim();
                            if (key && val) staticData[key] = val;
                        }
                    }
                }
                
                // Vessel name from title
                let vesselName = null;
                let titleElem = document.querySelector("h1.title");
                if (titleElem) {
                    vesselName = titleElem.innerText.replace("current position", "").trim();
                }
                
                // Build result object
                return {
                    static: {
                        vessel_name: staticData["vessel name"] || staticData["name"] || vesselName,
                        imo: (staticData["imo / mmsi"] ? staticData["imo / mmsi"].split('/')[0].trim() : staticData["imo"]) || null,
                        mmsi: (staticData["imo / mmsi"] ? staticData["imo / mmsi"].split('/')[1].trim() : staticData["mmsi"]) || null,
                        callsign: staticData["callsign"] || null,
                        ais_type: staticData["ais type"] || null,
                        flag: staticData["ais flag"] || staticData["flag"] || null,
                        length_beam: staticData["length / beam"] || null
                    },
                    position: {
                        status: staticData["navigation status"] || staticData["status"] || null,
                        course_speed: staticData["course / speed"] || null,
                        current_draught: draught ? draught + " m" : staticData["current draught"] || null,
                        navigation_status: staticData["navigation status"] || null,
                        position_received: posAge || null,
                        latitude: lat,
                        longitude: lon,
                        last_port: lastPort || null,
                        atd: atd || null
                    },
                    eta: {
                        destination: dest || null,
                        eta: eta || null
                    }
                };
            }""")
            
            return data
        except Exception as e:
            print(f"Error scraping {mmsi_or_imo}: {e}")
            return None
        finally:
            await browser.close()

# ------------------------------------------------------------
# Main: read config, scrape each vessel, save results
# ------------------------------------------------------------
async def main():
    if not os.path.exists(CONFIG_PATH):
        print(f"Error: Config file not found at {CONFIG_PATH}")
        return
    
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    
    # Load existing data if any (to preserve previous data for vessels that fail)
    if os.path.exists(DATA_PATH):
        with open(DATA_PATH, "r", encoding="utf-8") as f:
            vessel_data = json.load(f)
    else:
        vessel_data = {}
    
    for vessel_name, mmsi in config.items():
        mmsi_str = str(mmsi)
        print(f"\n--- Processing {vessel_name} (MMSI {mmsi_str}) ---")
        data = await scrape_vessel_with_js(mmsi_str)
        if data:
            vessel_data[mmsi_str] = data
            print(f"✅ Successfully updated {vessel_name}")
        else:
            print(f"❌ Failed to update {vessel_name}")
            # Optionally keep old data if it exists; otherwise set a placeholder
            if mmsi_str not in vessel_data:
                vessel_data[mmsi_str] = {"error": "Scraping failed"}
        
        await human_delay(3, 6)  # be polite to the server
    
    # Save updated data
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(vessel_data, f, indent=4, ensure_ascii=False)
    print(f"\n✅ All tasks complete. Data saved to {DATA_PATH}")

if __name__ == "__main__":
    asyncio.run(main())