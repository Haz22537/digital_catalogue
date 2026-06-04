# -*- coding: utf-8 -*-
"""
Racing Australia — Phase 1: URL & Dam Discovery
================================================

Searches Racing Australia for each horse in the filtered RA manifest,
captures the profile URL and Dam name, and inline-checks the dam against
the manifest dam. Phase 2 only processes rows that auto-matched, so any
mismatches stay in your review queue without polluting downstream data.

Output file dam_match column
----------------------------
    auto      — RA dam matches manifest dam (case-insensitive). Phase 2 will process.
    mismatch  — RA returned a dam but it differs. Investigate before changing to 'auto'.
    none      — Could not extract a dam (search returned nothing, or scrape failed).

Workflow
--------
1. Run this script. It writes every row with dam_match populated.
2. Open the file in Excel, filter dam_match != 'auto'.
3. Fix those rows (replace ra_profile_url if wrong, then set dam_match='auto'),
   or leave them be to exclude from Phase 2.
4. Save and run Phase 2.

Usage
-----
    python 02_RA_url_discovery.py --sale 2026-05B

Inputs
------
    {sale}_manifest_RA.csv     (output of manifest_convert.py)

Outputs
-------
    {sale}_RA_urls.csv         (input for Phase 2)
    {sale}_failures.csv        (appended — any rows that errored)

Resume behaviour
----------------
Output is written incrementally and flushed after every row. If the script
dies, re-running picks up where it left off: any Lot already in the output
is skipped, and previously-resolved horse names are reused so duplicates
across lots don't re-hit RA. To retry a failed row, delete it from the
output CSV and re-run.
"""

import argparse
import csv
import re
import sys
import time
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from sale_codes import manifest_RA_path, ra_urls_path
from pipeline import log_failure, normalize_dam


# ── Config ──────────────────────────────────────────────────────────
INPUT_COL          = "Name"
DAM_COL            = "Dam"
OUTPUT_EXTRA_COLS  = ["ra_dam", "ra_profile_url", "dam_match"]
BASE_URL           = "https://racingaustralia.horse/home.aspx"
COURTESY_DELAY_S   = 2  # between successive RA requests
SALE_PHASE         = "ra_url_discovery"


# ── Helpers ─────────────────────────────────────────────────────────
def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip()).lower()


def classify_dam_match(manifest_dam: str, ra_dam: str | None) -> str:
    """Compare manifest dam vs RA-extracted dam → 'auto' / 'mismatch' / 'none'."""
    if not ra_dam:
        return "none"
    if normalize_dam(manifest_dam) == normalize_dam(ra_dam):
        return "auto"
    return "mismatch"


def extract_dam(driver) -> str | None:
    """Pull only the Dam from the Horse Full Form page. Cheap parse."""
    soup = BeautifulSoup(driver.page_source, "html.parser")
    h2 = soup.find("h2", class_="first")
    if not h2:
        return None

    full_text = h2.parent.get_text(" ", strip=True)
    dam_match = re.search(r"\bfrom\s+(.+)$", full_text, re.IGNORECASE)
    if not dam_match:
        return None

    dam = dam_match.group(1).strip()
    dam = dam.split("View Pedigree")[0].strip()
    return dam or None


def load_resume_state(output_path: str) -> tuple[set[str], dict[str, dict]]:
    """Read existing output and return (lots_done, name_cache).

    name_cache only includes successful lookups so failed/empty rows get a retry.
    """
    lots_done: set[str] = set()
    name_cache: dict[str, dict] = {}

    if not Path(output_path).exists():
        return lots_done, name_cache

    with open(output_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            lot = (row.get("Lot") or "").strip()
            if lot:
                lots_done.add(lot)
            name = (row.get(INPUT_COL) or "").strip()
            url  = (row.get("ra_profile_url") or "").strip()
            if name and url:
                name_cache[normalize_name(name)] = {
                    "ra_dam":         row.get("ra_dam", ""),
                    "ra_profile_url": url,
                }

    return lots_done, name_cache


def scrape_one(driver, horse_name_raw: str) -> dict:
    """Search RA, navigate to profile, return {ra_dam, ra_profile_url}."""
    result = {"ra_dam": None, "ra_profile_url": None}

    driver.get(BASE_URL)

    WebDriverWait(driver, 20).until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.CSS_SELECTOR, "iframe[src*='SearchBox.aspx']")
        )
    )
    container = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "form-searchbox-id-jq"))
    )

    search_box = container.find_element(By.NAME, "boxSearch")
    btn_search = container.find_element(By.ID, "btnSearch")
    search_box.clear()
    search_box.send_keys(horse_name_raw)
    btn_search.click()

    driver.switch_to.default_content()

    # Direct hit vs multi-result list. We always take first link on multi —
    # dam isn't visible in that listing, so we rely on the post-nav dam check.
    try:
        WebDriverWait(driver, 5).until(EC.url_contains("HorseFullForm.aspx"))
    except Exception:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#PedigreeResults tbody tr td a")
            )
        )
        driver.find_element(
            By.CSS_SELECTOR, "#PedigreeResults tbody tr td a"
        ).click()
        WebDriverWait(driver, 10).until(EC.url_contains("HorseFullForm.aspx"))

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "h2.first"))
    )

    result["ra_profile_url"] = driver.current_url
    result["ra_dam"]         = extract_dam(driver)
    return result


# ── Main ────────────────────────────────────────────────────────────
def main(sale: str) -> None:
    input_path  = manifest_RA_path(sale)
    output_path = ra_urls_path(sale)

    try:
        with open(input_path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        sys.exit(
            f"❌ Input not found: {input_path}\n"
            f"   Run manifest_convert.py --sale {sale} first."
        )

    if not rows:
        sys.exit(f"❌ No rows in {input_path}")
    if INPUT_COL not in rows[0]:
        sys.exit(
            f"❌ Column '{INPUT_COL}' not found in {input_path}. "
            f"Available: {list(rows[0].keys())}"
        )
    if DAM_COL not in rows[0]:
        print(f"⚠  Column '{DAM_COL}' not found in input — dam_match will be 'none' for all rows.")

    output_cols = list(rows[0].keys()) + [
        c for c in OUTPUT_EXTRA_COLS if c not in rows[0]
    ]

    lots_done, name_cache = load_resume_state(output_path)
    resuming = bool(lots_done)
    if resuming:
        print(
            f"📦 Resuming: {len(lots_done)} lots already in {output_path}, "
            f"{len(name_cache)} cached lookups."
        )

    mode = "a" if resuming else "w"
    out_f = open(output_path, mode, newline="", encoding="utf-8")
    writer = csv.DictWriter(out_f, fieldnames=output_cols, extrasaction="ignore")
    if not resuming:
        writer.writeheader()
        out_f.flush()

    counts = {
        "scraped": 0, "cached": 0, "skipped": 0, "failed": 0,
        "auto": 0, "mismatch": 0, "none": 0,
    }

    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1280,900")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    try:
        for row in rows:
            lot          = (row.get("Lot") or "").strip()
            horse_raw    = (row.get(INPUT_COL) or "").strip()
            manifest_dam = (row.get(DAM_COL) or "").strip()

            if not horse_raw or horse_raw.lower() == "nan":
                continue
            if lot and lot in lots_done:
                counts["skipped"] += 1
                continue

            key = normalize_name(horse_raw)

            if key in name_cache:
                result = dict(name_cache[key])
                print(f"♻️  Cached: {horse_raw}")
                counts["cached"] += 1
            else:
                print(f"🔍 Searching: {horse_raw}")
                try:
                    result = scrape_one(driver, horse_raw)
                    name_cache[key] = dict(result)
                    counts["scraped"] += 1
                    time.sleep(COURTESY_DELAY_S)
                except Exception as e:
                    err = str(e).splitlines()[0][:200]
                    print(f"⚠️  Failed: {horse_raw} — {err}")
                    result = {"ra_dam": None, "ra_profile_url": None}
                    counts["failed"] += 1
                    log_failure(
                        sale, SALE_PHASE, lot=lot, horse_name=horse_raw,
                        status="scrape_error", fail_reason=err,
                    )

            match = classify_dam_match(manifest_dam, result.get("ra_dam"))
            counts[match] += 1

            if match == "mismatch":
                print(
                    f"   ↳ dam mismatch: manifest={manifest_dam!r} "
                    f"vs RA={result.get('ra_dam')!r}"
                )

            out_row = {**row, **result, "dam_match": match}
            writer.writerow(out_row)
            out_f.flush()
            if lot:
                lots_done.add(lot)

    finally:
        driver.quit()
        out_f.close()

    print(
        f"\n✅ Phase 1 complete — {output_path}\n"
        f"   scraped: {counts['scraped']}  cached: {counts['cached']}  "
        f"skipped: {counts['skipped']}  failed: {counts['failed']}\n"
        f"   dam_match → auto: {counts['auto']}  "
        f"mismatch: {counts['mismatch']}  none: {counts['none']}\n"
        f"   Next: open {output_path}, filter dam_match != 'auto', review {counts['mismatch'] + counts['none']} rows."
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Phase 1: search Racing Australia, capture URL + dam, auto-flag matches."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    args = parser.parse_args()
    main(args.sale)
