"""
Tracxn Company Data Scraper
============================
Scrapes company intelligence data from tracxn.com for Indian apparel brands.
Outputs one Excel file with all brands as rows.

Usage:
    python tracxn_scraper.py

Credentials:
    Set TRACXN_EMAIL and TRACXN_PASSWORD environment variables, OR
    edit TRACXN_EMAIL / TRACXN_PASSWORD constants below.

Output columns (per brand):
    Brand, Tracxn URL, Founded, HQ, Website, Description, Business Model,
    Stage, Total Funding, Last Round Type, Last Round Amount, Last Round Date,
    Valuation, Revenue, Employees, Founders, Key Executives, Investors,
    No. of Investors, Competitors, Tags, Scraped At, Notes
"""

import os
import re
import time
import json
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
    TimeoutException, NoSuchElementException, WebDriverException
)
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# CONFIG — override via env vars or edit directly
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

PAGE_LOAD_TIMEOUT = 30   # seconds
ELEMENT_WAIT      = 12   # seconds
BETWEEN_BRANDS    = (4, 8)   # random sleep range (seconds) between brands

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
# DRIVER SETUP
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


def _safe_text(driver, by, selector, fallback=""):
    try:
        return _clean(driver.find_element(by, selector).text)
    except NoSuchElementException:
        return fallback


def _safe_texts(driver, by, selector):
    try:
        return [_clean(el.text) for el in driver.find_elements(by, selector) if _clean(el.text)]
    except Exception:
        return []


def _wait_for(driver, by, selector, timeout=ELEMENT_WAIT):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
    except TimeoutException:
        return None


def _dismiss_popup(driver):
    """Close cookie / login-wall popups if present."""
    for sel in [
        "button[aria-label='Close']",
        "button.close",
        "[class*='modal'] button[class*='close']",
        "[class*='popup'] button[class*='close']",
        "button[data-dismiss='modal']",
    ]:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            btn.click()
            time.sleep(0.5)
            return
        except Exception:
            pass


# =========================================================
# LOGIN
# =========================================================
def login(driver):
    if not TRACXN_EMAIL or not TRACXN_PASSWORD:
        print("[WARN] No Tracxn credentials provided. Some data may be gated.")
        return False

    print(f"[INFO] Logging in as {TRACXN_EMAIL} ...")
    try:
        driver.get(LOGIN_URL)
        _wait_for(driver, By.CSS_SELECTOR, "input[type='email']", timeout=15)

        driver.find_element(By.CSS_SELECTOR, "input[type='email']").send_keys(TRACXN_EMAIL)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(TRACXN_PASSWORD)

        # Submit
        submit = driver.find_element(
            By.CSS_SELECTOR,
            "button[type='submit'], input[type='submit'], button.login-btn"
        )
        submit.click()

        # Wait for redirect away from login page
        WebDriverWait(driver, 20).until(EC.url_changes(LOGIN_URL))
        print("[INFO] Login successful.")
        return True
    except Exception as e:
        print(f"[WARN] Login failed: {e}")
        return False


# =========================================================
# SEARCH  →  return profile URL for first matching result
# =========================================================
def search_brand(driver, brand_name):
    query   = brand_name.replace(" ", "+")
    url     = SEARCH_URL.format(query=query)

    try:
        driver.get(url)
    except WebDriverException as e:
        print(f"    [WARN] Page load error during search: {e}")

    time.sleep(random.uniform(2, 4))
    _dismiss_popup(driver)

    # Try multiple selector patterns Tracxn uses
    result_selectors = [
        "a[href*='/d/companies/']",
        "[class*='company-name'] a",
        "[class*='companyName'] a",
        "h3 a[href*='/d/']",
        ".search-result a",
        "[data-type='company'] a",
    ]

    for sel in result_selectors:
        try:
            links = driver.find_elements(By.CSS_SELECTOR, sel)
            if links:
                href = links[0].get_attribute("href") or ""
                if href and "tracxn.com" in href:
                    return href
        except Exception:
            pass

    # Fallback: look for any tracxn company link on the page
    try:
        all_links = driver.find_elements(By.TAG_NAME, "a")
        for a in all_links:
            href = a.get_attribute("href") or ""
            if "/d/companies/" in href:
                return href
    except Exception:
        pass

    return ""


# =========================================================
# EXTRACT DATA FROM A COMPANY PROFILE PAGE
# =========================================================
def _extract_section_items(driver, section_keyword):
    """Return text items from a named section (e.g. 'Investors', 'Founders')."""
    items = []
    try:
        # Find section header containing the keyword
        headers = driver.find_elements(
            By.XPATH,
            f"//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{section_keyword.upper()}')]"
        )
        for h in headers:
            # Walk up to the section container and get list items
            try:
                container = h.find_element(By.XPATH, "./ancestor::*[.//*[contains(@class,'item') or contains(@class,'name') or contains(@class,'tag')]][1]")
                elems = container.find_elements(By.CSS_SELECTOR, "li, [class*='item'], [class*='name'], [class*='tag']")
                for el in elems:
                    t = _clean(el.text)
                    if t and t.lower() not in section_keyword.lower():
                        items.append(t)
                if items:
                    break
            except Exception:
                pass
    except Exception:
        pass
    return items


def _try_selectors(driver, *selectors):
    for sel in selectors:
        try:
            if sel.startswith("//") or sel.startswith("(//"):
                el = driver.find_element(By.XPATH, sel)
            else:
                el = driver.find_element(By.CSS_SELECTOR, sel)
            t = _clean(el.text)
            if t:
                return t
        except Exception:
            pass
    return ""


def _try_selectors_multi(driver, *selectors):
    for sel in selectors:
        try:
            if sel.startswith("//") or sel.startswith("(//"):
                els = driver.find_elements(By.XPATH, sel)
            else:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
            texts = [_clean(el.text) for el in els if _clean(el.text)]
            if texts:
                return texts
        except Exception:
            pass
    return []


def extract_company_data(driver, brand_name, profile_url):
    record = {
        "Brand":              brand_name,
        "Tracxn URL":         profile_url,
        "Founded":            "",
        "HQ":                 "",
        "Website":            "",
        "Description":        "",
        "Business Model":     "",
        "Stage":              "",
        "Total Funding":      "",
        "Last Round Type":    "",
        "Last Round Amount":  "",
        "Last Round Date":    "",
        "Valuation":          "",
        "Revenue":            "",
        "Employees":          "",
        "Founders":           "",
        "Key Executives":     "",
        "Investors":          "",
        "No. of Investors":   "",
        "Competitors":        "",
        "Tags":               "",
        "Scraped At":         datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Notes":              "",
    }

    if not profile_url:
        record["Notes"] = "Not found on Tracxn"
        return record

    try:
        driver.get(profile_url)
    except WebDriverException as e:
        record["Notes"] = f"Page load error: {e}"
        return record

    time.sleep(random.uniform(3, 5))
    _dismiss_popup(driver)

    # ---- Description / About ----
    record["Description"] = _try_selectors(
        driver,
        "[class*='description']",
        "[class*='about'] p",
        "[class*='companyDesc']",
        "//h2[contains(text(),'About')]/following-sibling::*[1]",
        "//h3[contains(text(),'About')]/following-sibling::*[1]",
        "meta[name='description']",
    )
    # meta fallback
    if not record["Description"]:
        try:
            meta = driver.find_element(By.CSS_SELECTOR, "meta[name='description']")
            record["Description"] = _clean(meta.get_attribute("content") or "")
        except Exception:
            pass

    # ---- Founded / HQ ----
    record["Founded"] = _try_selectors(
        driver,
        "[class*='founded']",
        "[class*='Founded']",
        "//span[contains(text(),'Founded')]/following-sibling::span[1]",
        "//td[contains(text(),'Founded')]/following-sibling::td[1]",
        "//div[contains(text(),'Founded')]/following-sibling::div[1]",
        "//th[contains(text(),'Founded')]/following-sibling::th[1]",
    )

    record["HQ"] = _try_selectors(
        driver,
        "[class*='location']",
        "[class*='headquarters']",
        "[class*='hq']",
        "//span[contains(text(),'Headquarter')]/following-sibling::span[1]",
        "//td[contains(text(),'Location')]/following-sibling::td[1]",
        "//div[contains(text(),'Location')]/following-sibling::div[1]",
    )

    # ---- Website ----
    record["Website"] = _try_selectors(
        driver,
        "a[class*='website']",
        "[class*='websiteUrl'] a",
        "//a[contains(@href,'http') and not(contains(@href,'tracxn'))]",
    )
    if not record["Website"]:
        try:
            links = driver.find_elements(By.CSS_SELECTOR, "a[href^='http']")
            for a in links:
                href = a.get_attribute("href") or ""
                if "tracxn" not in href and "linkedin" not in href and "twitter" not in href:
                    record["Website"] = href
                    break
        except Exception:
            pass

    # ---- Business Model / Stage ----
    record["Business Model"] = _try_selectors(
        driver,
        "[class*='businessModel']",
        "[class*='business-model']",
        "//span[contains(text(),'Business Model')]/following-sibling::span[1]",
        "//td[contains(text(),'Business Model')]/following-sibling::td[1]",
    )

    record["Stage"] = _try_selectors(
        driver,
        "[class*='stage']",
        "[class*='fundingStage']",
        "//span[contains(text(),'Stage')]/following-sibling::span[1]",
        "//td[contains(text(),'Stage')]/following-sibling::td[1]",
    )

    # ---- Funding ----
    record["Total Funding"] = _try_selectors(
        driver,
        "[class*='totalFunding']",
        "[class*='total-funding']",
        "//span[contains(text(),'Total Funding')]/following-sibling::span[1]",
        "//td[contains(text(),'Total Funding')]/following-sibling::td[1]",
        "//div[contains(text(),'Total Funding')]/following-sibling::div[1]",
        "//h4[contains(text(),'Total Funding')]/following-sibling::*[1]",
    )

    record["Last Round Type"] = _try_selectors(
        driver,
        "[class*='lastRoundType']",
        "[class*='roundType']",
        "//span[contains(text(),'Last Round')]/following-sibling::span[1]",
        "//td[contains(text(),'Latest Round')]/following-sibling::td[1]",
    )

    record["Last Round Amount"] = _try_selectors(
        driver,
        "[class*='lastRoundAmount']",
        "[class*='roundAmount']",
        "//td[contains(text(),'Amount')]/following-sibling::td[1]",
    )

    record["Last Round Date"] = _try_selectors(
        driver,
        "[class*='lastRoundDate']",
        "[class*='roundDate']",
        "//td[contains(text(),'Date')]/following-sibling::td[1]",
    )

    # ---- Valuation / Revenue / Employees ----
    record["Valuation"] = _try_selectors(
        driver,
        "[class*='valuation']",
        "//span[contains(text(),'Valuation')]/following-sibling::span[1]",
        "//td[contains(text(),'Valuation')]/following-sibling::td[1]",
    )

    record["Revenue"] = _try_selectors(
        driver,
        "[class*='revenue']",
        "//span[contains(text(),'Revenue')]/following-sibling::span[1]",
        "//td[contains(text(),'Revenue')]/following-sibling::td[1]",
    )

    record["Employees"] = _try_selectors(
        driver,
        "[class*='employees']",
        "[class*='teamSize']",
        "//span[contains(text(),'Employee')]/following-sibling::span[1]",
        "//td[contains(text(),'Employee')]/following-sibling::td[1]",
        "//span[contains(text(),'Team Size')]/following-sibling::span[1]",
    )

    # ---- People ----
    founders = _try_selectors_multi(
        driver,
        "[class*='founder'] [class*='name']",
        "[class*='Founder'] [class*='name']",
        "//h2[contains(text(),'Founder')]/following-sibling::*//span[contains(@class,'name')]",
        "//h3[contains(text(),'Founder')]/following-sibling::*//span[contains(@class,'name')]",
    )
    record["Founders"] = "; ".join(founders[:10])

    executives = _try_selectors_multi(
        driver,
        "[class*='executive'] [class*='name']",
        "[class*='team'] [class*='name']",
        "[class*='keyPeople'] [class*='name']",
    )
    record["Key Executives"] = "; ".join(executives[:10])

    # ---- Investors ----
    investors = _try_selectors_multi(
        driver,
        "[class*='investor'] [class*='name']",
        "[class*='investors'] li",
        "[class*='investorName']",
        "//h2[contains(text(),'Investor')]/following-sibling::*//a",
        "//h3[contains(text(),'Investor')]/following-sibling::*//a",
    )
    record["Investors"] = "; ".join(investors[:20])
    record["No. of Investors"] = str(len(investors)) if investors else ""

    # ---- Competitors ----
    competitors = _try_selectors_multi(
        driver,
        "[class*='competitor'] [class*='name']",
        "[class*='competitors'] li",
        "[class*='similarCompanies'] [class*='name']",
        "//h2[contains(text(),'Competitor')]/following-sibling::*//a",
        "//h3[contains(text(),'Competitor')]/following-sibling::*//a",
    )
    record["Competitors"] = "; ".join(competitors[:15])

    # ---- Tags / Categories ----
    tags = _try_selectors_multi(
        driver,
        "[class*='tag']",
        "[class*='category']",
        "[class*='sector']",
        "[class*='vertical']",
        "//span[contains(@class,'tag')]",
    )
    # Deduplicate short tags (typically < 50 chars)
    tags_clean = list(dict.fromkeys(t for t in tags if len(t) < 60))
    record["Tags"] = "; ".join(tags_clean[:15])

    # ---- Fallback: scrape visible page text for missed fields ----
    _fill_from_page_text(driver, record)

    return record


def _fill_from_page_text(driver, record):
    """Parse raw page text as a last-resort to fill missing fields."""
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return

    patterns = {
        "Founded":       r'(?:Founded|Incorporated)[:\s]+(\d{4})',
        "Employees":     r'(\d[\d,]+)\s+(?:employees|Employees|team members)',
        "Total Funding": r'Total Funding[:\s]+([^\n]{2,30})',
        "Revenue":       r'Revenue[:\s]+([^\n]{2,30})',
        "Valuation":     r'Valuation[:\s]+([^\n]{2,30})',
        "HQ":            r'(?:Headquarters|HQ|Location)[:\s]+([^\n]{2,60})',
        "Stage":         r'Stage[:\s]+([^\n]{2,40})',
    }

    for field, pattern in patterns.items():
        if not record[field]:
            m = re.search(pattern, body_text, re.IGNORECASE)
            if m:
                record[field] = _clean(m.group(1))


# =========================================================
# MAIN PIPELINE
# =========================================================
def run():
    today  = datetime.today().strftime("%Y-%m-%d")
    outdir = os.path.dirname(os.path.abspath(__file__))
    output = os.path.join(outdir, f"Tracxn_Brands_{today}.xlsx")

    print("=" * 70)
    print("  TRACXN BRAND INTELLIGENCE SCRAPER")
    print(f"  Brands  : {len(BRANDS)}")
    print(f"  Output  : {output}")
    print("=" * 70)

    driver = make_driver()
    logged_in = login(driver)
    if not logged_in:
        print("[INFO] Proceeding without login (public data only).")

    records = []

    for i, brand in enumerate(BRANDS, 1):
        print(f"\n[{i:>3}/{len(BRANDS)}] {brand}")

        # 1. Search for the brand
        print(f"    Searching ...")
        profile_url = search_brand(driver, brand)
        if profile_url:
            print(f"    Found  : {profile_url}")
        else:
            print(f"    [WARN] No result found for '{brand}'")

        # 2. Extract data from profile page
        record = extract_company_data(driver, brand, profile_url)
        records.append(record)

        # Progress summary
        filled = sum(1 for k, v in record.items()
                     if v and k not in ("Brand", "Tracxn URL", "Scraped At", "Notes"))
        print(f"    Fields : {filled} non-empty")

        # 3. Checkpoint save every 10 brands
        if i % 10 == 0:
            _save_excel(records, output)
            print(f"    [CHECKPOINT] Saved {i} brands so far.")

        # 4. Polite delay
        sleep_time = random.uniform(*BETWEEN_BRANDS)
        time.sleep(sleep_time)

    driver.quit()

    _save_excel(records, output)
    print("\n" + "=" * 70)
    print(f"  DONE. {len(records)} brands saved to:")
    print(f"  {output}")
    print("=" * 70)
    return output


def _save_excel(records, path):
    if not records:
        return
    df = pd.DataFrame(records)
    col_order = [
        "Brand", "Tracxn URL", "Founded", "HQ", "Website", "Description",
        "Business Model", "Stage", "Total Funding", "Last Round Type",
        "Last Round Amount", "Last Round Date", "Valuation", "Revenue",
        "Employees", "Founders", "Key Executives", "Investors",
        "No. of Investors", "Competitors", "Tags", "Scraped At", "Notes",
    ]
    df = df.reindex(columns=col_order)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tracxn Data")
        ws = writer.sheets["Tracxn Data"]

        # Column widths
        widths = {
            "A": 22, "B": 45, "C": 10, "D": 22, "E": 30, "F": 60,
            "G": 20, "H": 18, "I": 18, "J": 18, "K": 18, "L": 16,
            "M": 16, "N": 16, "O": 12, "P": 30, "Q": 30, "R": 50,
            "S": 10, "T": 45, "U": 35, "V": 18, "W": 35,
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

        # Freeze header row
        ws.freeze_panes = "A2"

    print(f"    [SAVE] {path}")


# =========================================================
# ENTRY POINT
# =========================================================
if __name__ == "__main__":
    run()
