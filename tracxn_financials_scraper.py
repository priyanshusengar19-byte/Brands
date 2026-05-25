"""
Tracxn Financials Scraper
==========================
Extracts "Revenue from sale of goods" for ALL available fiscal years
from the Tracxn Financials section for each brand.

IMPORTANT: Financial data on Tracxn requires a logged-in account.
Set credentials via env vars before running:
    export TRACXN_EMAIL="your@email.com"
    export TRACXN_PASSWORD="yourpassword"
    python tracxn_financials_scraper.py

Output: Tracxn_Financials_YYYY-MM-DD.xlsx
    Sheet 1 "Wide"  — brands as rows, each FY as a column
    Sheet 2 "Long"  — brand | Fiscal Year | Revenue from Sale of Goods (Cr)
"""

import os
import re
import time
import random
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException,
    ElementClickInterceptedException,
)
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# CONFIG
# =========================================================
TRACXN_EMAIL    = os.environ.get("TRACXN_EMAIL", "")
TRACXN_PASSWORD = os.environ.get("TRACXN_PASSWORD", "")

LOGIN_URL  = "https://tracxn.com/login"
SEARCH_URL = "https://tracxn.com/d/search?q={query}&type=companies"

_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

PAGE_LOAD_TIMEOUT = 30
ELEMENT_WAIT      = 15
BETWEEN_BRANDS    = (5, 10)

# Row label to look for in the income statement table
TARGET_ROW_LABEL = "Revenue from sale of goods"

BRANDS = [
    "Zudio",
    "Westside",
    "Louis Philippe",
    "Allen Solly",
    "Van Heusen",
    "Peter England",
    "Pantaloons",
    "Lifestyle",
    "Shoppers Stop",
    "Reliance Trends",
    "Azorte",
    "Yousta",
    "Max Fashion",
    "Style Union",
    "OWND",
    "Intune Fashion",
    "Vishal Mega Mart",
    "V2 Retail",
    "V-Mart",
    "Unlimited Fashion",
    "Style Bazaar",
    "Bazaar Kolkata",
    "CityBazaar Metro",
    "W for Woman",
    "Raymond",
    "Zodiac Clothing",
    "Blackberrys",
    "Rare Rabbit",
    "Rareism",
    "Nike India",
    "Reebok India",
    "Adidas India",
    "Puma India",
    "Fila India",
    "Mochi Shoes",
    "Metro Shoes",
    "Zara India",
    "H&M India",
    "United Colors of Benetton",
    "FabIndia",
    "Manyavar",
    "Mohey",
    "Relaxo Footwear",
    "Bata India",
    "Hush Puppies India",
    "Jockey India",
    "Go Colors",
    "US Polo Assn India",
    "Arrow India",
    "Calvin Klein India",
    "Tommy Hilfiger India",
    "Flying Machine",
    "Arvind Fashions",
    "Reliance Smart Bazaar",
    "MBazaar",
    "CityKart",
    "V-Baazar",
    "Levi's India",
    "Mufti",
    "Spykar",
    "Pepe Jeans India",
    "Jack and Jones India",
    "Biba India",
    "Aurelia",
    "Soch",
    "Marks and Spencer India",
    "ONLY India",
    "Park Avenue",
    "Lacoste India",
    "Red Tape",
    "Liberty Shoes",
    "Skechers India",
    "Khadims",
    "Cantabil",
    "Easy Buy",
    "Celio India",
    "Tasva",
    "Indian Terrain",
    "Monte Carlo",
    "Numero Uno",
    "Duke Fashion",
    "Siyaram's",
    "AND India",
    "Madame",
    "Kazo",
    "ColorPlus",
    "Parx",
    "John Player",
    "Lee Cooper India",
    "Killer Jeans",
    "Global Desi",
    "Woodland India",
    "Ritu Kumar",
    "Libas",
    "Oxemberg",
    "Snitch",
    "Powerlook",
    "The Souled Store",
    "Vero Moda India",
    "Asian Footwear",
    "Asics India",
    "Campus Shoes",
    "Crocs India",
    "Decathlon India",
    "Forever New",
    "Octave Apparel",
    "Turtle Ltd",
    "Wildcraft India",
    "Wrogn",
    "Bewakoof",
    "BonkersCorner",
    "Newme",
]

# =========================================================
# DRIVER
# =========================================================
def make_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument(f"--user-agent={_UA}")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts,
    )
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
    )
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver


# =========================================================
# HELPERS
# =========================================================
def _clean(text):
    return " ".join((text or "").replace("\n", " ").replace("\r", " ").split())


def _dismiss_popup(driver):
    for sel in [
        "button[aria-label='Close']",
        "button.close",
        "[class*='modal'] button[class*='close']",
        "[class*='popup'] button[class*='close']",
        "button[data-dismiss='modal']",
    ]:
        try:
            driver.find_element(By.CSS_SELECTOR, sel).click()
            time.sleep(0.4)
            return
        except Exception:
            pass


def _parse_cr_value(raw):
    """
    Convert Tracxn value strings to float (in Crores).
    Handles: '122.89Cr', '1,234.56Cr', '-', '', '0'
    Returns '' for missing/dash values.
    """
    raw = _clean(raw)
    if not raw or raw in ("-", "—", "N/A", "NA"):
        return ""
    raw = raw.replace(",", "").replace("Cr", "").strip()
    try:
        return float(raw)
    except ValueError:
        return raw  # return as-is if unparseable


# =========================================================
# LOGIN
# =========================================================
def login(driver):
    if not TRACXN_EMAIL or not TRACXN_PASSWORD:
        print("[WARN] No credentials — financial data will likely be blocked.")
        return False

    print(f"[INFO] Logging in as {TRACXN_EMAIL} ...")
    try:
        driver.get(LOGIN_URL)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
        )
        driver.find_element(By.CSS_SELECTOR, "input[type='email']").send_keys(TRACXN_EMAIL)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(TRACXN_PASSWORD)
        driver.find_element(
            By.CSS_SELECTOR,
            "button[type='submit'], input[type='submit']"
        ).click()
        WebDriverWait(driver, 20).until(EC.url_changes(LOGIN_URL))
        print("[INFO] Login successful.")
        return True
    except Exception as e:
        print(f"[WARN] Login failed: {e}")
        return False


# =========================================================
# SEARCH → profile URL
# =========================================================
def search_brand(driver, brand_name):
    url = SEARCH_URL.format(query=brand_name.replace(" ", "+"))
    try:
        driver.get(url)
    except WebDriverException:
        pass

    time.sleep(random.uniform(2, 3))
    _dismiss_popup(driver)

    for sel in [
        "a[href*='/d/companies/']",
        "[class*='companyName'] a",
        "[class*='company-name'] a",
        "h3 a[href*='/d/']",
    ]:
        try:
            links = driver.find_elements(By.CSS_SELECTOR, sel)
            if links:
                href = links[0].get_attribute("href") or ""
                if "tracxn.com" in href:
                    return href
        except Exception:
            pass

    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = a.get_attribute("href") or ""
            if "/d/companies/" in href:
                return href
    except Exception:
        pass

    return ""


# =========================================================
# NAVIGATE TO FINANCIALS TAB
# =========================================================
def go_to_financials(driver, profile_url):
    """
    Load the company profile and click the Financials sidebar link.
    Returns True if the financial table is visible.
    """
    try:
        driver.get(profile_url)
    except WebDriverException:
        return False

    time.sleep(random.uniform(2, 3))
    _dismiss_popup(driver)

    # Strategy 1: click sidebar "Financials" link
    fin_clicked = False
    for sel in [
        "//a[normalize-space(text())='Financials']",
        "//span[normalize-space(text())='Financials']/..",
        "//li[contains(@class,'nav')]//a[contains(text(),'Financials')]",
        "a[href*='financials']",
    ]:
        try:
            if sel.startswith("//"):
                el = driver.find_element(By.XPATH, sel)
            else:
                el = driver.find_element(By.CSS_SELECTOR, sel)
            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            el.click()
            fin_clicked = True
            break
        except Exception:
            pass

    # Strategy 2: direct URL append
    if not fin_clicked:
        slug_url = profile_url.rstrip("/") + "/financials"
        try:
            driver.get(slug_url)
        except WebDriverException:
            pass

    time.sleep(random.uniform(3, 5))
    _dismiss_popup(driver)

    # Confirm financials content loaded
    for marker in [
        "Detailed Income Statement",
        "Revenue from sale",
        "Total revenue",
        "Financial Type",
    ]:
        if marker.lower() in driver.page_source.lower():
            return True

    return False


# =========================================================
# EXPAND ALL FISCAL YEAR SECTIONS
# =========================================================
def expand_all_fy_sections(driver):
    """
    Click every collapsed FY section header so all years' data is visible.
    Also clicks 'Expand All' buttons in income statement tables.
    """
    # Expand collapsed FY year accordions (look for right-pointing arrows / chevrons)
    expanded = 0
    fy_header_selectors = [
        "//div[contains(@class,'fy') and contains(@class,'header')]",
        "//div[contains(@class,'year-header')]",
        "//div[contains(@class,'accordion')][.//text()[contains(.,'FY 20')]]",
        "//*[contains(@class,'collaps')][.//text()[contains(.,'FY 20')]]",
        # Generic: any element whose text looks like "FY YYYY-YY"
        "//*[re:match(text(), 'FY \\d{4}-\\d{2}')]",
    ]

    fy_pattern = re.compile(r'FY\s*\d{4}-\d{2}')

    # Find all elements containing FY year text
    all_els = driver.find_elements(By.XPATH, "//*[contains(text(),'FY 20')]")
    for el in all_els:
        txt = _clean(el.text)
        if not fy_pattern.match(txt):
            continue
        # Check if it looks like a section header (not a table cell)
        tag = el.tag_name.lower()
        if tag in ("td", "th", "span") and len(txt) < 15:
            # Might be a table header — click the parent section toggle
            try:
                parent = el.find_element(By.XPATH, "./ancestor::*[contains(@class,'accordion') or contains(@class,'collapse') or contains(@class,'section')][1]")
                aria = parent.get_attribute("aria-expanded") or ""
                if aria.lower() == "false":
                    parent.click()
                    expanded += 1
                    time.sleep(0.4)
            except Exception:
                pass
        elif tag in ("div", "h2", "h3", "h4", "button", "summary"):
            try:
                aria = el.get_attribute("aria-expanded") or ""
                cls  = el.get_attribute("class") or ""
                if aria.lower() == "false" or "collapsed" in cls:
                    el.click()
                    expanded += 1
                    time.sleep(0.4)
            except Exception:
                pass

    # Click any "Expand All [+]" buttons in income statement tables
    for btn_sel in [
        "//button[contains(text(),'Expand All')]",
        "//a[contains(text(),'Expand All')]",
        "//*[contains(text(),'Expand All')]",
        "//*[contains(text(),'[+]')]",
    ]:
        try:
            btns = driver.find_elements(By.XPATH, btn_sel)
            for b in btns:
                try:
                    b.click()
                    time.sleep(0.3)
                    expanded += 1
                except ElementClickInterceptedException:
                    pass
        except Exception:
            pass

    if expanded:
        time.sleep(1)  # let DOM settle after expansions

    return expanded


# =========================================================
# EXTRACT REVENUE FROM SALE OF GOODS — ALL FY
# =========================================================
FY_RE    = re.compile(r'FY\s*(\d{4}-\d{2})', re.IGNORECASE)
VALUE_RE = re.compile(r'-?\d[\d,]*\.?\d*\s*Cr', re.IGNORECASE)


def _extract_from_tables(driver):
    """
    Parse every <table> on the page.
    Returns dict: { 'FY 2024-25': 122.89, 'FY 2023-24': 98.5, ... }
    for the TARGET_ROW_LABEL row.
    """
    results = {}

    tables = driver.find_elements(By.TAG_NAME, "table")
    for table in tables:
        try:
            rows = table.find_elements(By.TAG_NAME, "tr")
            if not rows:
                continue

            # Parse header row to find FY column indices
            header_row = rows[0]
            headers = [_clean(th.text) for th in header_row.find_elements(By.XPATH, ".//th | .//td")]
            fy_col_map = {}  # col_index → FY label
            for ci, h in enumerate(headers):
                m = FY_RE.search(h)
                if m:
                    fy_col_map[ci] = f"FY {m.group(1)}"

            if not fy_col_map:
                continue

            # Scan data rows for the target label
            for row in rows[1:]:
                cells = row.find_elements(By.XPATH, ".//td | .//th")
                if not cells:
                    continue
                first_cell_text = _clean(cells[0].text)
                if TARGET_ROW_LABEL.lower() not in first_cell_text.lower():
                    continue
                # Found the target row — extract values per FY column
                for ci, fy_label in fy_col_map.items():
                    if ci < len(cells):
                        val = _parse_cr_value(cells[ci].text)
                        if val != "" and fy_label not in results:
                            results[fy_label] = val
        except Exception:
            pass

    return results


def _extract_from_page_structure(driver):
    """
    Fallback: walk the DOM looking for FY section headers, then find the
    target row within that section's subtree.
    Returns dict: { 'FY 2024-25': 122.89, ... }
    """
    results = {}

    # Find all FY section headers
    candidates = driver.find_elements(By.XPATH, "//*[contains(text(),'FY 20')]")
    fy_headers = []
    for el in candidates:
        txt = _clean(el.text)
        m = FY_RE.match(txt)
        if m:
            fy_headers.append((f"FY {m.group(1)}", el))

    for fy_label, header_el in fy_headers:
        try:
            # Walk up to the section container
            section = header_el.find_element(
                By.XPATH,
                "./ancestor::*[contains(@class,'section') or contains(@class,'accordion') "
                "or contains(@class,'card') or contains(@class,'panel') or self::section][1]"
            )
        except Exception:
            section = header_el

        # Search for the target row within this section
        try:
            target_cells = section.find_elements(
                By.XPATH,
                f".//*[contains(text(),'{TARGET_ROW_LABEL}')]"
            )
            for tc in target_cells:
                # Get the sibling/following cell with a number
                row_el = tc.find_element(By.XPATH, "./ancestor::tr[1]")
                cells  = row_el.find_elements(By.TAG_NAME, "td")
                for cell in cells:
                    raw = _clean(cell.text)
                    if VALUE_RE.search(raw):
                        val = _parse_cr_value(raw)
                        if val != "" and fy_label not in results:
                            results[fy_label] = val
                            break
        except Exception:
            pass

    return results


def _extract_from_raw_text(driver):
    """
    Last-resort: scan raw page text for patterns near 'Revenue from sale of goods'.
    Returns dict: { 'FY 2024-25': 122.89, ... }
    """
    results = {}
    try:
        body = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return results

    # Split into lines and look for the target label surrounded by FY context
    lines = body.splitlines()
    current_fy = None
    for i, line in enumerate(lines):
        fy_m = FY_RE.search(line)
        if fy_m:
            current_fy = f"FY {fy_m.group(1)}"

        if TARGET_ROW_LABEL.lower() in line.lower() and current_fy:
            # Value might be on the same line or the next line
            for candidate in [line] + (lines[i+1:i+3] if i + 1 < len(lines) else []):
                val_m = VALUE_RE.search(candidate)
                if val_m:
                    val = _parse_cr_value(val_m.group(0))
                    if val != "" and current_fy not in results:
                        results[current_fy] = val
                    break

    return results


def scrape_financials(driver, brand_name, profile_url):
    """
    Full pipeline: navigate → expand → extract for one brand.
    Returns dict: { 'FY 2024-25': 122.89, 'FY 2023-24': ..., ... }
    plus meta keys 'brand', 'tracxn_url', 'notes'.
    """
    meta = {
        "brand":       brand_name,
        "tracxn_url":  profile_url,
        "notes":       "",
    }

    if not profile_url:
        meta["notes"] = "Not found on Tracxn"
        return meta, {}

    fin_loaded = go_to_financials(driver, profile_url)
    if not fin_loaded:
        meta["notes"] = "Financials section not loaded (login required or page error)"
        return meta, {}

    expand_all_fy_sections(driver)

    # Try extraction strategies in order of reliability
    data = _extract_from_tables(driver)
    if not data:
        data = _extract_from_page_structure(driver)
    if not data:
        data = _extract_from_raw_text(driver)

    if not data:
        meta["notes"] = "No financial data found (may be gated)"

    return meta, data


# =========================================================
# MAIN
# =========================================================
def run():
    today  = datetime.today().strftime("%Y-%m-%d")
    outdir = os.path.dirname(os.path.abspath(__file__))
    output = os.path.join(outdir, f"Tracxn_Financials_{today}.xlsx")

    print("=" * 70)
    print("  TRACXN FINANCIALS SCRAPER — Revenue from Sale of Goods")
    print(f"  Brands : {len(BRANDS)}")
    print(f"  Output : {output}")
    print("=" * 70)

    driver    = make_driver()
    logged_in = login(driver)
    if not logged_in:
        print("[WARN] Not logged in — financial data is likely gated on Tracxn.")

    all_meta  = []   # list of meta dicts
    all_data  = []   # list of {brand, FY, value} for long format

    for i, brand in enumerate(BRANDS, 1):
        print(f"\n[{i:>3}/{len(BRANDS)}] {brand}")

        # Step 1: find profile URL
        profile_url = search_brand(driver, brand)
        print(f"    URL : {profile_url or '(not found)'}")

        # Step 2: scrape financials
        meta, fy_data = scrape_financials(driver, brand, profile_url)
        all_meta.append(meta)

        if fy_data:
            years_str = ", ".join(sorted(fy_data.keys(), reverse=True))
            print(f"    FY  : {years_str}")
            for fy, val in sorted(fy_data.items(), reverse=True):
                all_data.append({
                    "Brand":                       brand,
                    "Fiscal Year":                 fy,
                    "Revenue from Sale of Goods (Cr)": val,
                    "Tracxn URL":                  profile_url,
                })
        else:
            print(f"    [WARN] {meta['notes'] or 'No data'}")
            all_data.append({
                "Brand":                       brand,
                "Fiscal Year":                 "",
                "Revenue from Sale of Goods (Cr)": "",
                "Tracxn URL":                  profile_url,
            })

        # Checkpoint save every 10 brands
        if i % 10 == 0:
            _save_excel(all_meta, all_data, output)
            print(f"    [CHECKPOINT] Saved {i} brands.")

        time.sleep(random.uniform(*BETWEEN_BRANDS))

    driver.quit()

    _save_excel(all_meta, all_data, output)
    print("\n" + "=" * 70)
    print(f"  DONE. Results saved to:\n  {output}")
    print("=" * 70)


# =========================================================
# EXCEL WRITER — wide + long sheets
# =========================================================
def _save_excel(all_meta, all_data, path):
    if not all_data:
        return

    # ---------- Long sheet ----------
    long_df = pd.DataFrame(all_data, columns=[
        "Brand", "Fiscal Year", "Revenue from Sale of Goods (Cr)", "Tracxn URL"
    ])
    long_df = long_df[long_df["Fiscal Year"] != ""]  # drop placeholder rows

    # ---------- Wide sheet ----------
    if not long_df.empty:
        wide_df = long_df.pivot_table(
            index=["Brand", "Tracxn URL"],
            columns="Fiscal Year",
            values="Revenue from Sale of Goods (Cr)",
            aggfunc="first",
        ).reset_index()
        # Sort FY columns newest → oldest
        fy_cols = sorted(
            [c for c in wide_df.columns if FY_RE.match(str(c))],
            reverse=True
        )
        wide_df = wide_df[["Brand", "Tracxn URL"] + fy_cols]
    else:
        wide_df = pd.DataFrame(columns=["Brand", "Tracxn URL"])

    # ---------- Notes sheet ----------
    notes_df = pd.DataFrame([
        {"Brand": m["brand"], "Tracxn URL": m["tracxn_url"], "Notes": m["notes"]}
        for m in all_meta if m.get("notes")
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Wide sheet
        wide_df.to_excel(writer, index=False, sheet_name="Wide (by FY)")
        _fmt_sheet(writer.sheets["Wide (by FY)"], wide_df)

        # Long sheet
        long_df.to_excel(writer, index=False, sheet_name="Long (all records)")
        _fmt_sheet(writer.sheets["Long (all records)"], long_df)

        # Notes
        if not notes_df.empty:
            notes_df.to_excel(writer, index=False, sheet_name="Notes")
            _fmt_sheet(writer.sheets["Notes"], notes_df)

    print(f"    [SAVE] {path}")


def _fmt_sheet(ws, df):
    ws.freeze_panes = "A2"
    for i, col in enumerate(df.columns, 1):
        max_len = max(
            len(str(col)),
            *(len(str(v)) for v in df[col].fillna("") if v != ""),
            default=10,
        )
        ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max_len + 4, 50)


# =========================================================
# ENTRY POINT
# =========================================================
if __name__ == "__main__":
    run()
