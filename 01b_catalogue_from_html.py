# -*- coding: utf-8 -*-
"""
Inglis Digital Catalogue → Manifest, from a saved HTML file
=============================================================

Same output as 01_catalogue_scrape.py, but skips Playwright/browser
automation entirely. Takes a pre-saved copy of the rendered catalogue
table (View Source → Save As, or a manual copy of the table container's
outerHTML) and parses it directly with BeautifulSoup.

Why this exists
----------------
01_catalogue_scrape.py drives a real browser to force the Vue SPA to
render, then scrolls until the virtualised list is fully loaded, then
walks the DOM live via Playwright locators (one round-trip per cell).
That's slow and has a known timeout on data-col-key under load.

Once the page is scrolled to the bottom in a real browser (so every lot
has actually rendered), the resulting HTML is static and self-contained
— there is nothing left for a browser to do. This script just parses
that saved HTML directly, which is near-instant and has no timeout risk.

How to get the input file
--------------------------
1. Open the catalogue page in a normal browser.
2. Scroll to the very bottom so all lots have rendered (Inglis lazy-loads
   the table).
3. Right-click the table → Inspect → find the table container → right
   click the element → Copy → Copy outerHTML. Paste into a .txt/.html file.
   (Full "View Source" also works but is a much bigger file — either is fine,
   this script only looks for <td data-col-key="..."> cells.)

Usage
-----
    python 01b_catalogue_from_html.py --sale 2026-07A --input 2026-07A_html.txt

Outputs
-------
    {sale}_manifest.csv   (same shape as 01_catalogue_scrape.py's output —
                            input for inglisdigital_lotscraper.py)
"""

import argparse
import csv
import re
import sys

from bs4 import BeautifulSoup

from sale_codes import manifest_path
from pipeline import log_failure


SALE_PHASE = "catalogue_html"

# Output column shape — matches inglisdigital_lotscraper.py's expectations.
MANIFEST_COLUMNS = [
    "Lot", "Age", "Cat", "Sex", "Name",
    "Sire", "Dam", "Vendor", "Covering Stallion", "State",
]

# Inglis catalogue cell key → our manifest column.
# Same mapping as 01_catalogue_scrape.py — keep these in sync if Inglis
# change the catalogue markup.
COL_KEY_MAP = {
    "lotNumber":                    "Lot",
    "animalBirthDate":              "Age",   # Pre-computed e.g. "3YO", "W"
    "categoryCode":                 "Cat",
    "animalGenderCode":             "Sex",
    "animalName":                   "Name",  # Special: extract from <a> only
    "sireName":                     "Sire",
    "damName":                      "Dam",
    "vendorDisplayNameInCatalogue": "Vendor",
    "coveringStallionName":         "Covering Stallion",
    "horseLocationState":           "State",
}


def extract_rows(soup: BeautifulSoup) -> list[dict]:
    """Walk every <tr> in the parsed HTML and pull each lot's fields."""
    rows_out = []

    for tr in soup.find_all("tr"):
        cells = tr.find_all("td", attrs={"data-col-key": True})
        if not cells:
            continue  # header row, sticky-header clone, etc.

        row: dict[str, str] = {}
        for cell in cells:
            key = cell.get("data-col-key", "")
            if key not in COL_KEY_MAP:
                continue

            if key == "animalName":
                # Cell holds "<a>Name (AUS)</a><span>mobile-only Sire x Dam</span>".
                # Only the <a> text is the real name.
                anchor = cell.find("a")
                value = anchor.get_text(strip=True) if anchor else cell.get_text(strip=True)
            else:
                value = cell.get_text(strip=True)

            row[COL_KEY_MAP[key]] = re.sub(r"\s+", " ", value)

        if not row.get("Lot") or not row.get("Name"):
            continue

        for col in MANIFEST_COLUMNS:
            row.setdefault(col, "")

        rows_out.append(row)

    # Deduplicate by Lot (the live table has a fixed sticky-header clone of
    # row 1 plus the real body — belt-and-braces here too)
    seen = set()
    deduped = []
    for r in rows_out:
        if r["Lot"] in seen:
            continue
        seen.add(r["Lot"])
        deduped.append(r)

    try:
        deduped.sort(key=lambda r: int(r["Lot"]))
    except ValueError:
        deduped.sort(key=lambda r: r["Lot"])

    return deduped


def run(sale: str, input_path: str) -> None:
    out_path = manifest_path(sale)

    print(f"\n📂 Input:  {input_path}")
    print(f"   Output: {out_path}\n")

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            html = f.read()
    except OSError as e:
        log_failure(sale, SALE_PHASE, status="input_read_failed", fail_reason=str(e))
        sys.exit(f"❌ Could not read {input_path}: {e}")

    soup = BeautifulSoup(html, "html.parser")
    rows = extract_rows(soup)

    if not rows:
        log_failure(sale, SALE_PHASE, status="no_lots_found",
                    fail_reason="Parsed HTML but extracted 0 rows — "
                                "col_key map may be stale or file is incomplete")
        sys.exit(
            "❌ No rows extracted. Check the saved HTML actually contains "
            "the rendered table body (scroll to bottom before saving), and "
            "that data-col-key names still match COL_KEY_MAP."
        )

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in MANIFEST_COLUMNS})

    print(f"✅ Wrote {len(rows)} lots → {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Parse a saved Inglis Digital catalogue HTML file → manifest CSV."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-07A")
    parser.add_argument("--input", required=True, help="Path to saved catalogue HTML/txt file")
    args = parser.parse_args()
    run(args.sale, args.input)
