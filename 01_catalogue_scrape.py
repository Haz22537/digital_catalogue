# -*- coding: utf-8 -*-
"""
Inglis Digital Catalogue → Manifest Scraper
============================================

Scrapes the public catalogue page for a sale and writes the initial manifest
CSV. Replaces the hand-built manifest step.

How the page is structured (verified against a real DOM dump)
-------------------------------------------------------------
The catalogue is a Vue.js SPA. After it renders, the lot list is a single
HTML table where every <td> carries a `data-col-key` attribute identifying
the field, and an inner element carries a `data-entry-id` that is constant
across all cells for one lot. We:
    1. Wait for the table to fully render.
    2. Walk every <tr>.
    3. For each row, read its cells by data-col-key into a dict.
    4. Map to manifest columns and write.

This is far more stable than href-walking — Inglis can re-skin the page
without breaking the column names.

Usage
-----
    python 01_catalogue_scrape.py --sale 2026-05B

Diagnostic mode (only needed if Inglis change the markup):
    python 01_catalogue_scrape.py --sale 2026-05B --snapshot

Snapshot mode dumps the rendered HTML + a screenshot so the data-col-key
mapping can be re-verified offline.

Outputs
-------
    {sale}_manifest.csv   (input for inglisdigital_lotscraper.py)
"""

import argparse
import asyncio
import csv
import re
import sys

from playwright.async_api import async_playwright

from sale_codes import (
    short_to_slug, manifest_path, catalogue_snapshot_path,
)
from pipeline import log_failure


BASE_URL = "https://www.inglisdigital.com"
SALE_PHASE = "catalogue"

# Output column shape — matches what inglisdigital_lotscraper.py expects.
MANIFEST_COLUMNS = [
    "Lot", "Age", "Cat", "Sex", "Name",
    "Sire", "Dam", "Vendor", "Covering Stallion", "State",
]

# Inglis catalogue cell key → our manifest column.
# Keys verified against a captured DOM dump of 2026-05B catalogue page.
# If extraction breaks, run --snapshot and confirm these keys still appear.
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

REQUIRED_KEYS = {"lotNumber", "animalName"}


# ── Page interaction ──────────────────────────────────────────────────

async def wait_for_table(page) -> None:
    """Wait for the catalogue table cells to render."""
    await page.wait_for_selector("td[data-col-key='lotNumber']", timeout=30000)


async def load_all_lots(page) -> int:
    """Scroll until the row count is stable.

    Some Inglis catalogues virtualise the list (only render visible rows),
    so we scroll progressively until two consecutive checks return the same
    count. Returns the final row count.
    """
    prev = -1
    stable = 0
    for _ in range(60):
        count = await page.locator("td[data-col-key='lotNumber']").count()
        if count == prev:
            stable += 1
            if stable >= 3:
                break
        else:
            stable = 0
        prev = count
        await page.mouse.wheel(0, 3000)
        await page.wait_for_timeout(400)
    return prev


async def take_snapshot(page, sale: str) -> None:
    html_path = catalogue_snapshot_path(sale)
    png_path  = html_path.replace(".html", ".png")
    html = await page.content()
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    await page.screenshot(path=png_path, full_page=True)
    print(f"   📸 Wrote {html_path} + {png_path}")


# ── Per-row extraction ────────────────────────────────────────────────

async def extract_rows(page) -> list[dict]:
    """Walk every <tr> in the catalogue and pull each lot's fields.

    Returns one dict per lot keyed by manifest column names.
    """
    rows_out = []
    trs = await page.locator("tr").all()

    for tr in trs:
        cells = await tr.locator("td[data-col-key]").all()
        if not cells:
            continue  # header row, separator, etc.

        row: dict[str, str] = {}
        for cell in cells:
            key = await cell.get_attribute("data-col-key") or ""
            if key not in COL_KEY_MAP:
                continue

            if key == "animalName":
                # The cell contains "<a>Name (AUS)</a><span>mobile-only Sire x Dam</span>".
                # Only the <a> text is the real name.
                anchor = cell.locator("a").first
                if await anchor.count():
                    value = (await anchor.inner_text()).strip()
                else:
                    value = (await cell.inner_text()).strip()
            else:
                value = (await cell.inner_text()).strip()

            row[COL_KEY_MAP[key]] = re.sub(r"\s+", " ", value)

        # Skip rows that didn't yield the required identity fields.
        if not row.get("Lot") or not row.get("Name"):
            continue

        # Fill any missing manifest columns with blank
        for col in MANIFEST_COLUMNS:
            row.setdefault(col, "")

        rows_out.append(row)

    # Deduplicate by Lot (defensive — virtualised lists can re-render rows)
    seen = set()
    deduped = []
    for r in rows_out:
        if r["Lot"] in seen:
            continue
        seen.add(r["Lot"])
        deduped.append(r)

    # Sort numerically by Lot
    try:
        deduped.sort(key=lambda r: int(r["Lot"]))
    except ValueError:
        deduped.sort(key=lambda r: r["Lot"])

    return deduped


# ── Main ──────────────────────────────────────────────────────────────

async def run(sale: str, snapshot: bool) -> None:
    slug          = short_to_slug(sale)
    catalogue_url = f"{BASE_URL}/catalogue/auction/{slug}"
    out_path      = manifest_path(sale)

    print(f"\n📂 Catalogue: {catalogue_url}")
    print(f"   Output:    {out_path}\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1600, "height": 900},
        )
        page = await context.new_page()

        try:
            await page.goto(catalogue_url, timeout=30000)
            await page.wait_for_load_state("networkidle", timeout=20000)
            await wait_for_table(page)
        except Exception as e:
            log_failure(sale, SALE_PHASE, status="page_load_failed", fail_reason=str(e))
            await browser.close()
            sys.exit(f"❌ Failed to load catalogue: {e}")

        final_count = await load_all_lots(page)
        print(f"   Catalogue settled at {final_count} lot cells.")

        if snapshot:
            await take_snapshot(page, sale)
            await browser.close()
            print("\n✅ Snapshot complete.\n")
            return

        rows = await extract_rows(page)
        await browser.close()

    if not rows:
        log_failure(sale, SALE_PHASE, status="no_lots_found",
                    fail_reason="Found cells but extracted 0 rows — col_key map may be stale")
        sys.exit(
            "❌ No rows extracted. Re-run with --snapshot and send "
            f"{catalogue_snapshot_path(sale)} so the col_key map can be re-verified."
        )

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in MANIFEST_COLUMNS})

    print(f"\n✅ Wrote {len(rows)} lots → {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Scrape Inglis Digital catalogue page → initial manifest CSV."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    parser.add_argument(
        "--snapshot", action="store_true",
        help="Dump page HTML + screenshot, no CSV. For selector debugging.",
    )
    args = parser.parse_args()
    asyncio.run(run(args.sale, args.snapshot))
