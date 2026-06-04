# -*- coding: utf-8 -*-
"""
Racing Australia — Phase 2: Ratings & Trainer Enrichment
=========================================================

Reads Phase 1 output, transforms each profile URL to a ratings-page URL,
and scrapes the latest rating row and trainer details. Plain requests, no
Selenium.

Gating
------
Only rows with dam_match == 'auto' are processed. Rows with 'mismatch' or
'none' are skipped with a console warning and logged to the failure file
so they're visible during triage. To process a previously-mismatched row,
edit Phase 1's output, set dam_match to 'auto', and re-run.

NSW rating split
----------------
Racing NSW reports two benchmark ratings separated by a comma:
    "75,62" → Metro/Provincial=75, Country=62
    "0,62"  → didn't race Metro/Prov, Country=62
Other jurisdictions report a single value. Split logic:
    - Comma present  → left value into Rating, right value into Rating_NSW_Country
    - No comma       → value stays in Rating; Rating_NSW_Country left blank
A literal '0' is preserved (real data: 'raced 0 times here'); blank means
'no NSW data applies'.

Usage
-----
    python 03_RA_ratings_trainer.py --sale 2026-05B

Inputs
------
    {sale}_RA_urls.csv         (output of 02_RA_url_discovery.py, edited if needed)

Outputs
-------
    {sale}_RA_ratings.csv      (final RA-enriched data)
    {sale}_failures.csv        (appended)

Resume behaviour
----------------
Output is written and flushed incrementally after every row — the same
pattern as Phase 1. If the script dies mid-run, rows already written are
preserved. Re-running will reuse any row whose Lot already has a rating or
trainer in the output file. To retry a failed row, delete it from the
output CSV and re-run.
"""

import argparse
import csv
import re
import sys
import time
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

import requests
from bs4 import BeautifulSoup

from sale_codes import ra_urls_path, ra_ratings_path
from pipeline import log_failure


RATINGS_BASE = "https://racingaustralia.horse/FreeServices/HorseSearch_Ratings_Flats.aspx"
COURTESY_DELAY_S = 0.3
SALE_PHASE = "ra_ratings"

TITLE_PATTERN = re.compile(
    r"^\s*(Mr|Mrs|Ms|Miss|Dr|Prof|Rev)\.?\s+", re.IGNORECASE
)

NEW_COLS = [
    "ra_rating_url", "trainer_name", "trainer_location",
    "Pos", "Date", "Track", "Dist", "Class", "Rating", "Rating_NSW_Country", "Rating Type",
]


# ── Helpers ─────────────────────────────────────────────────────────
def build_ratings_url(profile_url: str) -> str | None:
    """Extract HorseCode from profile URL and build the ratings page URL."""
    try:
        params = parse_qs(urlparse(profile_url).query)
        horse_code = params.get("HorseCode", [None])[0]
        if not horse_code:
            return None
        return f"{RATINGS_BASE}?{urlencode({'HorseCode': horse_code})}"
    except Exception:
        return None


def split_nsw_rating(raw_rating: str) -> tuple[str, str]:
    """Split a NSW two-benchmark rating into (metro_prov, country).

    Non-NSW horses (no comma) → (raw, "").
    Blank input → ("", "").
    """
    if not raw_rating:
        return "", ""
    if "," in raw_rating:
        left, _, right = raw_rating.partition(",")
        return left.strip(), right.strip()
    return raw_rating.strip(), ""


def extract_trainer(soup: BeautifulSoup) -> tuple[str | None, str | None]:
    """Return (name, location). Name has any honorific stripped."""
    for th in soup.find_all("th"):
        if th.get_text(strip=True).lower() == "trainer":
            td = th.find_next_sibling("td")
            if not td:
                return None, None

            a = td.find("a", class_="GreenLink")
            raw_name = a.get_text(strip=True) if a else None
            name = TITLE_PATTERN.sub("", raw_name).strip() if raw_name else None

            full_text = td.get_text(" ", strip=True)
            loc_match = re.search(r"\(([^)]+)\)", full_text)
            location = loc_match.group(1).strip() if loc_match else None
            return name, location
    return None, None


def extract_latest_rating(soup: BeautifulSoup) -> dict:
    """Scrape the last row of the ratings table — keyed by column header."""
    table = soup.select_one("table.horse-search-strip-fields")
    if not table:
        return {}
    rows = table.find_all("tr")
    if len(rows) <= 1:
        return {}
    headers = [th.get_text(strip=True) for th in rows[0].find_all("th")]
    values  = [td.get_text(strip=True) for td in rows[-1].find_all("td")]
    return dict(zip(headers, values))


# ── Main ────────────────────────────────────────────────────────────
def main(sale: str) -> None:
    input_path  = ra_urls_path(sale)
    output_path = ra_ratings_path(sale)

    try:
        with open(input_path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        sys.exit(
            f"❌ Input not found: {input_path}\n"
            f"   Run 02_RA_url_discovery.py --sale {sale} first."
        )

    if not rows:
        sys.exit(f"❌ No rows in {input_path}")
    if "ra_profile_url" not in rows[0]:
        sys.exit(
            f"❌ 'ra_profile_url' column missing in {input_path}. "
            f"Available: {list(rows[0].keys())}"
        )
    if "dam_match" not in rows[0]:
        sys.exit(
            f"❌ 'dam_match' column missing in {input_path}. "
            f"Re-run Phase 1 — this script requires the dam_match gate."
        )

    output_cols = list(rows[0].keys()) + [c for c in NEW_COLS if c not in rows[0]]

    # ── Resume: load any rows already written ──────────────────────
    prior: dict[str, dict] = {}
    if Path(output_path).exists():
        with open(output_path, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                lot = (r.get("Lot") or "").strip()
                if lot and (r.get("trainer_name") or r.get("Rating")):
                    prior[lot] = r
        if prior:
            print(f"📦 Resuming: {len(prior)} rows already enriched in {output_path}")

    counts = {
        "fetched": 0, "cached": 0, "failed": 0,
        "skipped_no_url": 0, "skipped_dam_match": 0,
    }

    # ── Open output file for incremental writing ───────────────────
    # Append mode when resuming so already-written rows are preserved.
    # Write mode (with header) when starting fresh.
    resuming = bool(prior)
    file_mode = "a" if resuming else "w"

    with open(output_path, file_mode, newline="", encoding="utf-8") as out_f:
        writer = csv.DictWriter(out_f, fieldnames=output_cols, extrasaction="ignore")
        if not resuming:
            writer.writeheader()
            out_f.flush()

        total = len(rows)

        for row in rows:
            lot         = (row.get("Lot") or "").strip()
            profile_url = (row.get("ra_profile_url") or "").strip()
            dam_match   = (row.get("dam_match") or "").strip().lower()
            horse_name  = (row.get("Name") or "").strip()

            # ── Already done in a prior run ────────────────────────
            if lot and lot in prior:
                # Row already in the file from a previous run — skip writing again.
                counts["cached"] += 1
                print(f"♻️  [{counts['cached']+counts['fetched']+counts['failed']}/{total}] Cached: {horse_name}")
                continue

            # ── Gate: only auto-matched rows proceed ───────────────
            if dam_match != "auto":
                counts["skipped_dam_match"] += 1
                log_failure(
                    sale, SALE_PHASE, lot=lot, horse_name=horse_name,
                    status=f"skipped_dam_match_{dam_match or 'blank'}",
                    fail_reason="Phase 2 only processes dam_match='auto' rows",
                )
                out_row = {**row, **{c: "" for c in NEW_COLS}}
                writer.writerow(out_row)
                out_f.flush()
                continue

            if not profile_url:
                counts["skipped_no_url"] += 1
                out_row = {**row, **{c: "" for c in NEW_COLS}}
                writer.writerow(out_row)
                out_f.flush()
                continue

            ratings_url = build_ratings_url(profile_url)
            if not ratings_url:
                log_failure(
                    sale, SALE_PHASE, lot=lot, horse_name=horse_name,
                    status="no_horse_code",
                    fail_reason=f"Could not parse HorseCode from {profile_url}",
                )
                counts["failed"] += 1
                out_row = {**row, **{c: "" for c in NEW_COLS}}
                writer.writerow(out_row)
                out_f.flush()
                continue

            enriched = {**row, "ra_rating_url": ratings_url}
            for c in NEW_COLS:
                enriched.setdefault(c, "")

            done_so_far = counts["fetched"] + counts["cached"] + counts["failed"]
            print(f"🔍 [{done_so_far + 1}/{total}] {horse_name or ratings_url}")

            try:
                r = requests.get(ratings_url, timeout=20)
                soup = BeautifulSoup(r.text, "html.parser")

                trainer_name, trainer_location = extract_trainer(soup)
                enriched["trainer_name"]     = trainer_name or ""
                enriched["trainer_location"] = trainer_location or ""

                latest = extract_latest_rating(soup)
                for k, v in latest.items():
                    if k in NEW_COLS:
                        enriched[k] = v

                # NSW split: applies after raw scrape, only mutates if comma present
                raw_rating = enriched.get("Rating", "")
                metro_prov, country = split_nsw_rating(raw_rating)
                enriched["Rating"]             = metro_prov
                enriched["Rating_NSW_Country"] = country

                counts["fetched"] += 1

            except Exception as e:
                err = str(e).splitlines()[0][:200]
                print(f"   ⚠️  Error: {err}")
                counts["failed"] += 1
                log_failure(
                    sale, SALE_PHASE, lot=lot, horse_name=horse_name,
                    status="fetch_or_parse_error", fail_reason=err,
                )

            # ── Write and flush immediately after every row ────────
            writer.writerow(enriched)
            out_f.flush()
            time.sleep(COURTESY_DELAY_S)

    print(
        f"\n✅ Phase 2 complete — {output_path}\n"
        f"   fetched: {counts['fetched']}  cached: {counts['cached']}  "
        f"failed: {counts['failed']}\n"
        f"   skipped (no URL): {counts['skipped_no_url']}  "
        f"skipped (dam_match!=auto): {counts['skipped_dam_match']}"
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Phase 2: scrape ratings + trainer, NSW dual-rating split."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    args = parser.parse_args()
    main(args.sale)
