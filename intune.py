import re, time, random, json, requests
import pandas as pd
from bs4 import BeautifulSoup
import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# =========================================================
# CONFIG
# =========================================================
_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
URL = "https://intunefashion.com/pages/store-locator"

HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://intunefashion.com/",
    "User-Agent": _UA,
}

PINCODE_RE = re.compile(r'\b(\d{6})\b')
EMAIL_RE   = re.compile(r'Email\s*ID\s*[-:]\s*(\S+@\S+)', re.I)


# =========================================================
# HELPERS
# =========================================================
def _clean(text):
    return " ".join((text or "").replace("\n", " ").replace("\r", " ").replace("\xa0", " ").split())


def _get(session, retries=3):
    for attempt in range(1, retries + 1):
        try:
            r = session.get(URL, headers=HEADERS, timeout=30)
            if r.status_code == 200:
                return r
            if r.status_code in (429, 503, 502):
                time.sleep(2 ** attempt)
                continue
            return r
        except requests.RequestException:
            time.sleep(2 ** attempt)
    return None


# =========================================================
# SCRAPER
# =========================================================
def scrape_all(output_path):
    session = requests.Session()
    r = _get(session)
    if not r:
        print("[ERROR] Failed to fetch page")
        return pd.DataFrame()

    soup = BeautifulSoup(r.text, "html.parser")

    # Collect lat/lng from map marker inputs (positional order matches accordion order)
    coords = []
    for inp in soup.select("input.thb-location[data-option]"):
        try:
            d = json.loads(inp["data-option"])
            coords.append((str(d.get("latitude", "")), str(d.get("longitude", ""))))
        except Exception:
            coords.append(("", ""))

    # Parse accordion rows
    rows = []
    for idx, row in enumerate(soup.select("collapsible-row")):
        summary = row.find("summary")
        content = row.select_one("div.collapsible__content")
        if not summary or not content:
            continue

        # Store location name from <summary>
        location = _clean(summary.get_text())

        # City = last comma-segment of summary (e.g. "GSM Mall, Hyderabad" → "Hyderabad")
        city = location.split(",")[-1].strip().title() if "," in location else ""

        # Full address text — strip leading "Intune, " prefix if present
        raw_lines = [_clean(s) for s in content.get_text("\n").splitlines() if _clean(s)]

        # Separate zone label, address and email
        zone    = ""
        address = ""
        email   = ""
        addr_parts = []
        for line in raw_lines:
            em = EMAIL_RE.search(line)
            if em:
                email = em.group(1)
                continue
            if re.search(r'\bZone\b', line, re.I) and len(line) < 30:
                zone = line.rstrip(",").strip()
                continue
            addr_parts.append(line)

        address = " ".join(addr_parts).strip().rstrip(".")
        # Strip leading "Intune, " from address
        address = re.sub(r'^Intune,\s*', '', address, flags=re.I)

        # Pincode
        pin_m   = PINCODE_RE.search(address)
        pincode = pin_m.group(1) if pin_m else ""

        # State from address: word before pincode
        state = ""
        if pincode:
            before = address[:address.index(pincode)].rstrip(", ")
            parts  = [p.strip() for p in before.split(",") if p.strip()]
            if parts:
                state = parts[-1].title()

        # Maps URL from positional lat/lng
        maps_url = ""
        if idx < len(coords) and coords[idx][0]:
            lat, lng = coords[idx]
            maps_url = f"https://www.google.com/maps?q={lat},{lng}"

        rows.append({
            "Store Name": "Intune",
            "Location":   location,
            "City":       city,
            "State":      state,
            "Zone":       zone,
            "Address":    address,
            "Pincode":    pincode,
            "Email":      email,
            "Store Link": URL,
            "Maps URL":   maps_url,
        })

    df = pd.DataFrame(rows)
    col_order = ["Store Name", "Location", "City", "State", "Zone",
                 "Address", "Pincode", "Email", "Store Link", "Maps URL"]
    df = df.reindex(columns=col_order)

    df.to_excel(output_path, index=False)
    print(f"[INFO] Saved {len(df)} stores to {output_path}")
    return df


# =========================================================
# GUI
# =========================================================
def run_all():
    folder      = os.path.dirname(os.path.abspath(__file__))
    today       = datetime.today().strftime("%Y-%m-%d")
    output_path = os.path.join(folder, f"Intune_Stores_{today}.xlsx")

    print("=" * 70)
    print("[INFO] STARTING INTUNE SCRAPER")
    print("=" * 70)

    scrape_all(output_path)

    print("=" * 70)
    print("[INFO] FILE SAVED SUCCESSFULLY")
    print("=" * 70)

    messagebox.showinfo("Success", f"Saved to:\n{output_path}")


def run_gui():
    root = tk.Tk()
    root.title("Intune Store Scraper")
    root.geometry("430x220")
    root.resizable(False, False)

    tk.Label(
        root,
        text="Intune Store Scraper",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=20)

    tk.Button(
        root,
        text="Run",
        width=18,
        height=2,
        font=("Segoe UI", 12, "bold"),
        command=run_all
    ).pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
