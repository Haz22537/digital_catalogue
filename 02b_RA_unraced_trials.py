# -*- coding: utf-8 -*-
"""
Racing Australia — Phase 1b: Unraced Horse Trials & Trainer Scrape
===================================================================

Searches Racing Australia for each unraced horse in the filtered manifest,
captures the profile URL, dam name (for identity verification), trainer
name, trainer location (stable), and the most recent form entry — whether
that is a trial, jump-out, or race.

This script mirrors the navigation pattern of 02_RA_url_discovery.py but
extracts additional data from the HorseFullForm.aspx page: all the
information needed to populate the UNRACED tab in the Buyers' Guide is
available on this single page, so no second-page navigation is required.

Output columns added
--------------------
    ra_dam             — dam name as recorded on RA (used for identity check)
    ra_profile_url     — HorseFullForm.aspx URL
    dam_match          — 'auto' / 'mismatch' / 'none' (same logic as Phase 1)
    trainer_name       — trainer name, honorific stripped
    trainer_location   — trainer base location, e.g. 'Flemington'
    last_run_result    — finish position of most recent form entry, e.g. '3rd of 7'
    last_run_type      — 'J/O' (jumpout), 'TRIAL', or blank (race / no form)
    last_run_track     — track code, e.g. 'FLEM'
    last_run_date      — date of most recent run, DD/MM/YYYY

No form entries
---------------
Some unraced horses will have no form at all on RA — either genuinely
untried or too new to have any record. These rows are written with blank
trial fields. This is expected and accepted; do not treat blank as an error.

dam_match gate
--------------
Same as Phase 1: only 'auto' rows should be trusted for downstream use.
Mismatches and 'none' rows are written to the output for your review.
There is no Phase 2 for this pipeline — the HorseFullForm page contains
everything needed, so this is the only RA script run for unraced horses.

Manual review gate
------------------
After this script finishes, open {sale}_RA_unraced_trials.csv in Excel,
filter dam_match != 'auto', and review. Fix the URL if the wrong horse
was returned, or leave to exclude from the report build.

Usage
-----
    python 02b_RA_unraced_trials.py --sale 2026-05B

Inputs
------
    {sale}_manifest_RA_unraced.csv   (output of manifest_convert.py)

Outputs
-------
    {sale}_RA_unraced_trials.csv     (input for report build — UNRACED tab)
    {sale}_failures.csv              (appended — any rows that errored)

Resume behaviour
----------------
Output is written and flushed after every row. Re-running skips any Lot
already in the output file. Horse name cache avoids re-hitting RA for
duplicate names. To retry a failed row, delete it from the output CSV
and re-run.
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

from sale_codes import manifest_RA_unraced_path, ra_unraced_trials_path
from pipeline import log_failure, normalize_dam


# ── Config ──────────────────────────────────────────────────────────
INPUT_COL         = "Name"
DAM_COL           = "Dam"
BASE_URL          = "https://racingaustralia.horse/home.aspx"
COURTESY_DELAY_S  = 2
SALE_PHASE        = "ra_unraced_trials"

OUTPUT_EXTRA_COLS = [
    "ra_dam", "ra_profile_url", "dam_match",
    "trainer_name", "trainer_location",
    "last_run_result", "last_run_type", "last_run_track", "last_run_date",
]

# Title pattern to strip from trainer names (mirrors 03_RA_ratings_trainer.py)
TITLE_PATTERN = re.compile(
    r"^\s*(Mr|Mrs|Ms|Miss|Dr|Prof|Rev)\.?\s+", re.IGNORECASE
)

# Month abbreviation map for parsing RA date strings like '29May26'
MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}


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


def parse_ra_date(date_str: str) -> str | None:
    """
    Parse RA date strings like '29May26' or '18Dec25' → 'DD/MM/YYYY'.
    Returns None if the string doesn't match the expected pattern.
    """
    m = re.match(r"(\d{1,2})([A-Za-z]{3})(\d{2})$", date_str.strip())
    if not m:
        return None
    day, mon, yr = m.group(1), m.group(2).lower(), m.group(3)
    month_num = MONTH_MAP.get(mon)
    if not month_num:
        return None
    year_full = f"20{yr}"
    return f"{int(day):02d}/{month_num}/{year_full}"


def extract_page_data(driver) -> dict:
    """
    Extract all required fields from the HorseFullForm.aspx page currently
    loaded in the driver. Returns a dict with keys matching OUTPUT_EXTRA_COLS
    (minus ra_profile_url and dam_match which are set by the caller).

    Returns blank strings for all fields if the page cannot be parsed.
    Fields with no data (no trainer, no form) are returned as empty strings —
    not None — so CSV writing is clean.
    """
    result = {
        "ra_dam":           None,   # None signals extract failure to caller
        "trainer_name":     "",
        "trainer_location": "",
        "last_run_result":  "",
        "last_run_type":    "",
        "last_run_track":   "",
        "last_run_date":    "",
    }

    soup = BeautifulSoup(driver.page_source, "html.parser")

    # ── Dam (identity check) ─────────────────────────────────────────
    h2 = soup.find("h2", class_="first")
    if h2:
        full_text = h2.parent.get_text(" ", strip=True)
        dam_match = re.search(r"\bfrom\s+(.+)$", full_text, re.IGNORECASE)
        if dam_match:
            dam = dam_match.group(1).strip()
            dam = dam.split("View Pedigree")[0].strip()
            result["ra_dam"] = dam or None

    # ── Trainer name + location ──────────────────────────────────────
    # Structure: <th>Trainer</th><td><a class="GreenLink">Name</a> (Location)</td>
    for th in soup.find_all("th"):
        if th.get_text(strip=True).lower() == "trainer":
            td = th.find_next_sibling("td")
            if td:
                a = td.find("a", class_="GreenLink")
                if a:
                    raw_name = a.get_text(strip=True)
                    result["trainer_name"] = TITLE_PATTERN.sub("", raw_name).strip()
                full_text_td = td.get_text(" ", strip=True)
                loc_match = re.search(r"\(([^)]+)\)", full_text_td)
                if loc_match:
                    result["trainer_location"] = loc_match.group(1).strip()
            break

    # ── Most recent form entry (last row of the form table) ──────────
    # Table class: 'interactive-race-fields'
    # Each row has two cells: td.Pos (result) and td.remain (race details).
    # Trials/jumpouts have their Pos cell content wrapped in <i><b>J</b>...</i>
    # or <i><b>T</b>...</i>. Plain races have no italic wrapper.
    # We take the last row unconditionally — this catches any race run since
    # the Inglis race record was last updated, which would show as a blank type.
    form_table = soup.find("table", class_="interactive-race-fields")
    if form_table:
        data_rows = [
            tr for tr in form_table.find_all("tr")
            if tr.find("td", class_="Pos")
        ]
        if data_rows:
            last_row = data_rows[-1]

            # ── Result and type ──────────────────────────────────────
            pos_cell = last_row.find("td", class_="Pos")
            if pos_cell:
                italic = pos_cell.find("i")
                if italic:
                    # Trial or jumpout — extract the type letter and clean result
                    bold = italic.find("b")
                    type_letter = bold.get_text(strip=True).upper() if bold else ""
                    if type_letter == "J":
                        result["last_run_type"] = "J/O"
                    elif type_letter == "T":
                        result["last_run_type"] = "TRIAL"
                    else:
                        result["last_run_type"] = type_letter  # fallback
                    # Result text: full italic text minus the type letter prefix
                    full_pos = italic.get_text(" ", strip=True)
                    # Strip leading single letter + whitespace (e.g. "J 3rd of 7" → "3rd of 7")
                    clean_pos = re.sub(r"^[A-Z]\s+", "", full_pos).strip()
                    result["last_run_result"] = clean_pos
                else:
                    # Plain race — no italic wrapper, no type letter
                    result["last_run_type"]   = ""
                    result["last_run_result"] = pos_cell.get_text(" ", strip=True).strip()

            # ── Track and date ───────────────────────────────────────
            # The 'remain' cell contains a link like:
            #   <a class="GreenLink" href="Meeting.aspx?...">FLEM 29May26</a>
            # Track code = first token, date = second token.
            remain_cell = last_row.find("td", class_="remain")
            if remain_cell:
                meeting_link = remain_cell.find("a", class_="GreenLink")
                if meeting_link:
                    meeting_text = meeting_link.get_text(strip=True)
                    # Expected format: 'FLEM 29May26' or 'SEYM 15Jan26'
                    parts = meeting_text.split()
                    if len(parts) >= 2:
                        result["last_run_track"] = parts[0].upper()
                        parsed_date = parse_ra_date(parts[1])
                        result["last_run_date"] = parsed_date or parts[1]

    return result


def load_resume_state(output_path: str) -> tuple[set[str], dict[str, dict]]:
    """Read existing output; return (lots_done, name_cache).
    name_cache only includes successful lookups — failed rows get a retry.
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
                name_cache[normalize_name(name)] = {k: row.get(k, "") for k in OUTPUT_EXTRA_COLS}

    return lots_done, name_cache


def scrape_one(driver, horse_name_raw: str) -> dict:
    """Search RA, navigate to HorseFullForm.aspx, extract all fields."""
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

    # Direct hit vs multi-result list — same pattern as 02_RA_url_discovery.py
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

    profile_url = driver.current_url
    page_data   = extract_page_data(driver)
    page_data["ra_profile_url"] = profile_url
    return page_data


# ── Main ────────────────────────────────────────────────────────────
def main(sale: str) -> None:
    input_path  = manifest_RA_unraced_path(sale)
    output_path = ra_unraced_trials_path(sale)

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

    mode  = "a" if resuming else "w"
    out_f = open(output_path, mode, newline="", encoding="utf-8")
    writer = csv.DictWriter(out_f, fieldnames=output_cols, extrasaction="ignore")
    if not resuming:
        writer.writeheader()
        out_f.flush()

    counts = {
        "scraped": 0, "cached": 0, "skipped": 0, "failed": 0,
        "no_form": 0, "auto": 0, "mismatch": 0, "none": 0,
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
                page_data = dict(name_cache[key])
                print(f"♻️  Cached: {horse_raw}")
                counts["cached"] += 1
            else:
                print(f"🔍 Searching: {horse_raw}")
                try:
                    page_data = scrape_one(driver, horse_raw)
                    name_cache[key] = dict(page_data)
                    counts["scraped"] += 1
                    time.sleep(COURTESY_DELAY_S)
                except Exception as e:
                    err = str(e).splitlines()[0][:200]
                    print(f"⚠️  Failed: {horse_raw} — {err}")
                    page_data = {c: "" for c in OUTPUT_EXTRA_COLS}
                    page_data["ra_dam"] = None
                    counts["failed"] += 1
                    log_failure(
                        sale, SALE_PHASE, lot=lot, horse_name=horse_raw,
                        status="scrape_error", fail_reason=err,
                    )

            # ── Dam match classification ─────────────────────────────
            match = classify_dam_match(manifest_dam, page_data.get("ra_dam"))
            counts[match] += 1

            if match == "mismatch":
                print(
                    f"   ↳ dam mismatch: manifest={manifest_dam!r} "
                    f"vs RA={page_data.get('ra_dam')!r}"
                )

            # ── No-form indicator ────────────────────────────────────
            if not page_data.get("last_run_track"):
                counts["no_form"] += 1
                print(f"   ℹ️  No form on RA: {horse_raw} (blank trial fields written)")

            out_row = {**row, **page_data, "dam_match": match}
            writer.writerow(out_row)
            out_f.flush()
            if lot:
                lots_done.add(lot)

    finally:
        driver.quit()
        out_f.close()

    print(
        f"\n✅ Phase 1b complete — {output_path}\n"
        f"   scraped: {counts['scraped']}  cached: {counts['cached']}  "
        f"skipped: {counts['skipped']}  failed: {counts['failed']}\n"
        f"   no form on RA: {counts['no_form']} (blank trial fields — expected)\n"
        f"   dam_match → auto: {counts['auto']}  "
        f"mismatch: {counts['mismatch']}  none: {counts['none']}\n"
        f"   Next: review dam_match != 'auto' rows, then feed into report build."
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Phase 1b: scrape trainer + last trial/jumpout for unraced horses."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    args = parser.parse_args()
    main(args.sale)
