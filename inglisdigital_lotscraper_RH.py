"""
Racing Insights — Inglis Digital Catalogue Scraper
Outputs two CSVs per sale:
  {sale_short}_stats.csv  — structured/numeric fields, fast to filter and sort
  {sale_short}_text.csv   — heavy text fields (pedigree, race record, description)

Joined on: lot_id (sale_slug + lot number, e.g. "2026-march-early-online-sale_4")

Usage
-----
  Non-interactive (preferred):
      python inglisdigital_lotscraper.py --sale 2026-05B
      python inglisdigital_lotscraper.py --sale 2026-05B --start 30 --end 80

  Interactive (legacy fallback when --sale omitted):
      python inglisdigital_lotscraper.py

Requirements:
    pip install playwright
    playwright install chromium
"""

import argparse
import asyncio
import csv
import re
import sys
from datetime import datetime
from playwright.async_api import async_playwright

from sale_codes import (
    short_to_slug, slug_to_short,
    manifest_path, stats_path, text_path,
)
from pipeline import log_failure

BASE_URL = "https://www.inglisdigital.com"

# Only these categories are scraped — all others get a stub row only
SCRAPE_CATEGORIES = {"RH", "RF", "2YO"}


# ─────────────────────────────────────────────
# STATS CSV COLUMNS
# Structured fields — safe to filter, sort, and analyse directly.
# Same column order every run so files from multiple sales stack cleanly.
# ─────────────────────────────────────────────
STATS_FIELDS = [
    "lot_id",
    "sale_slug",
    "lot",
    "url",
    "scraped_at",
    # Identity
    "horse_name",
    "age",
    "sex",
    "category",
    "sire",
    "dam",
    "colour",
    "dob",
    "cob",
    "disclosures",
    "gst",
    "reserve_status",
    # Manifest-supplied fields
    "vendor",
    "covering_stallion",
    "state",
    # Race record — totals
    "total_runs",
    "total_wins",
    "total_2nd",
    "total_3rd",
    "total_earnings",
    # Race record — by age (2YO through 6YO covers most cases; extras appended dynamically)
    "age2_runs", "age2_wins", "age2_2nd", "age2_3rd", "age2_earnings",
    "age3_runs", "age3_wins", "age3_2nd", "age3_3rd", "age3_earnings",
    "age4_runs", "age4_wins", "age4_2nd", "age4_3rd", "age4_earnings",
    "age5_runs", "age5_wins", "age5_2nd", "age5_3rd", "age5_earnings",
    "age6_runs", "age6_wins", "age6_2nd", "age6_3rd", "age6_earnings",
    # Barrier trials
    "barrier_trial_count",
    # Reports
    "report_count",
    "reports",
    # Meta
    "race_record_updated",
    "pedigree_updated",
    "error",
]

# ─────────────────────────────────────────────
# TEXT CSV COLUMNS
# Heavy free-text fields. Joined to stats via lot_id.
# Dam columns (1st_dam, 2nd_dam…) are dynamic — appended as found.
# ─────────────────────────────────────────────
TEXT_FIELDS = [
    "lot_id",
    "sale_slug",
    "lot",
    "description",
    "sire_stats",
    "1st_dam",
    "2nd_dam",
    "3rd_dam",
    "4th_dam",
    "pedigree_race_record_summary",
    "race_record_full",
]


def read_manifest(path):
    """
    Read the lot manifest CSV → {lot_num (int): normalised_row_dict}.
    Expected columns (case-insensitive): Lot, Age, Cat, Sex, Name,
      Sire, Dam, Vendor, Covering Stallion, State
    """
    COL_MAP = {
        "lot":               "lot",
        "age":               "age",
        "cat":               "category",
        "sex":               "sex",
        "name":              "horse_name",
        "sire":              "sire",
        "dam":               "dam",
        "vendor":            "vendor",
        "covering stallion": "covering_stallion",
        "state":             "state",
    }
    manifest = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            norm = {
                COL_MAP.get(k.strip().lower(), k.strip().lower()): (v or "").strip()
                for k, v in row.items()
            }
            try:
                lot_num = int(norm.get("lot", ""))
            except (ValueError, TypeError):
                continue
            manifest[lot_num] = norm
    return manifest


# ─────────────────────────────────────────────
# ACCORDION HELPER
# ─────────────────────────────────────────────

async def open_accordion(page, accordion_id, wait_selector=None):
    """Click an accordion open and wait for content to render."""
    accordion = page.locator(f"#{accordion_id}")
    if await accordion.count() == 0:
        return False

    toggle = accordion.locator("div.cursor-pointer").first
    if await toggle.count() == 0:
        return False

    await toggle.evaluate("el => el.click()")

    if wait_selector:
        # Scope EACH comma-separated selector to #{accordion_id}. Without this,
        # only the first selector is scoped and any others match document-wide
        # — that caused wait_for_selector to time out at 6s on every lot when
        # passed ".arion-report, p.rem0", silently slowing the scrape.
        scoped = ", ".join(
            f"#{accordion_id} {part.strip()}"
            for part in wait_selector.split(",")
        )
        try:
            await page.wait_for_selector(scoped, timeout=6000)
        except Exception:
            pass
    else:
        await page.wait_for_timeout(600)

    return True


# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

def clean_whitespace(text):
    """Collapse excessive whitespace while keeping paragraph breaks."""
    # Normalise non-breaking spaces — Playwright's inner_text() returns \xa0
    # between paragraphs on the Inglis pedigree page; left in, they bloat
    # the CSV with invisible characters and trip up downstream text matching.
    text = text.replace("\xa0", " ")
    text = re.sub(r"\r\n|\r", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def make_lot_id(sale_slug, lot_number):
    return f"{sale_slug}_{lot_number}"


# Race record count columns — these display '-' on Inglis when the value is zero.
# Earnings columns are NOT in this set: zero earnings render as '$0.00', not '-'.
RACE_RECORD_COUNT_COLS = {
    "total_runs", "total_wins", "total_2nd", "total_3rd",
    "age2_runs", "age2_wins", "age2_2nd", "age2_3rd",
    "age3_runs", "age3_wins", "age3_2nd", "age3_3rd",
    "age4_runs", "age4_wins", "age4_2nd", "age4_3rd",
    "age5_runs", "age5_wins", "age5_2nd", "age5_3rd",
    "age6_runs", "age6_wins", "age6_2nd", "age6_3rd",
}


def zero_if_dash(value):
    """'-', '—', '–', '' → '0'. Used for race-record count columns."""
    if value is None:
        return "0"
    s = str(value).strip()
    if s in ("", "-", "—", "–"):
        return "0"
    return s


# ─────────────────────────────────────────────
# ACCORDION EXTRACTORS
# Each returns (stats_dict, text_dict) or just stats_dict for Reports.
# ─────────────────────────────────────────────

async def extract_overview(page):
    stats, text = {}, {}
    await open_accordion(page, "accordion_overview", wait_selector="p, table")
    accordion = page.locator("#accordion_overview")

    # Particulars table → stats
    try:
        rows = accordion.locator("tr")
        for i in range(await rows.count()):
            cells = rows.nth(i).locator("td")
            if await cells.count() >= 2:
                label = (await cells.nth(0).inner_text()).strip().lower()
                value = (await cells.nth(1).inner_text()).strip()
                key_map = {
                    "category": "category",
                    "sire": "sire",
                    "dam": "dam",
                    "colour": "colour",
                    "color": "colour",
                    "sex": "sex",
                    "dob": "dob",
                    "age": "age",
                    "cob": "cob",
                    "disclosures": "disclosures",
                    "gst": "gst",
                }
                for k, col in key_map.items():
                    if k in label:
                        stats[col] = value
    except Exception:
        pass

    # Description → text CSV
    try:
        full_text = await accordion.inner_text()
        desc_match = re.search(
            r"Description\s*\n+(.*?)(?:\n{3,}|\Z)", full_text, re.DOTALL
        )
        if desc_match:
            text["description"] = clean_whitespace(desc_match.group(1))
        else:
            parts = full_text.split("Description", 1)
            if len(parts) > 1:
                text["description"] = clean_whitespace(parts[1])
    except Exception:
        text["description"] = ""

    return stats, text


async def extract_pedigree(page):
    stats, text = {}, {}
    await open_accordion(
        page, "accordion_pedigree", wait_selector=".arion-report, p.rem0"
    )
    accordion = page.locator("#accordion_pedigree")

    try:
        full_text = await accordion.inner_text()
    except Exception:
        return stats, text

    # Last updated → stats
    updated_match = re.search(r"Pedigree Last Updated:\s*(.+)", full_text)
    if updated_match:
        stats["pedigree_updated"] = updated_match.group(1).strip()

    # Sire stats paragraph → text
    sire_match = re.search(
        r"([A-Z][A-Z &\(\)/]+\s*\((?:NZ|GB|AUS|IRE|USA)\).*?Stud \d{4}.*?etc\.)",
        full_text,
        re.DOTALL,
    )
    if sire_match:
        text["sire_stats"] = clean_whitespace(sire_match.group(1))

    # Dam narratives → text (1st dam, 2nd dam, 3rd dam, 4th dam).
    # NOTE: Playwright's inner_text() on the Inglis pedigree page returns
    # headings concatenated directly with content (e.g. "1st damOSTREIDAE,
    # by Pierro...") — no newline, no space between them. Paragraphs are
    # separated only by \xa0 (non-breaking space). The earlier `\s*\n`
    # requirement between heading and content therefore matched nothing and
    # silently produced empty dam fields for every lot. clean_whitespace()
    # normalises the \xa0 → space on the captured content.
    dam_sections = re.findall(
        r"(\d(?:st|nd|rd|th) dam)(.*?)(?=\d(?:st|nd|rd|th) dam|Race Record:|\Z)",
        full_text,
        re.DOTALL,
    )
    for label, content in dam_sections:
        key = label.replace(" ", "_").lower()  # "1st_dam", "2nd_dam", etc.
        text[key] = clean_whitespace(content)

    # Race record summary line at bottom of pedigree → text.
    # Same upstream issue as dam sections — no newline between the
    # "Race Record:" label and its content. End boundary is either a blank
    # line (where the Arion copyright block starts) or the literal
    # "All pedigrees" phrase that always follows on this page.
    rr_match = re.search(
        r"Race Record:\s*(.*?)(?=\n\s*\n|All pedigrees|\Z)",
        full_text,
        re.DOTALL,
    )
    if rr_match:
        text["pedigree_race_record_summary"] = clean_whitespace(rr_match.group(1))

    return stats, text


async def extract_race_record(page):
    stats, text = {}, {}
    await open_accordion(
        page, "accordion_raceRecord", wait_selector="table, p"
    )
    accordion = page.locator("#accordion_raceRecord")

    # Summary stats table → stats CSV
    try:
        rows = accordion.locator("table").first.locator("tr")
        for i in range(await rows.count()):
            cells = rows.nth(i).locator("td")
            cell_texts = [
                (await cells.nth(j).inner_text()).strip()
                for j in range(await cells.count())
            ]
            if not cell_texts:
                continue
            first = cell_texts[0]
            if first.isdigit():
                a = first
                stats[f"age{a}_runs"]     = zero_if_dash(cell_texts[1]) if len(cell_texts) > 1 else "0"
                stats[f"age{a}_wins"]     = zero_if_dash(cell_texts[2]) if len(cell_texts) > 2 else "0"
                stats[f"age{a}_2nd"]      = zero_if_dash(cell_texts[3]) if len(cell_texts) > 3 else "0"
                stats[f"age{a}_3rd"]      = zero_if_dash(cell_texts[4]) if len(cell_texts) > 4 else "0"
                stats[f"age{a}_earnings"] = cell_texts[5] if len(cell_texts) > 5 else ""
            elif "total" in first.lower():
                stats["total_runs"]     = zero_if_dash(cell_texts[1]) if len(cell_texts) > 1 else "0"
                stats["total_wins"]     = zero_if_dash(cell_texts[2]) if len(cell_texts) > 2 else "0"
                stats["total_2nd"]      = zero_if_dash(cell_texts[3]) if len(cell_texts) > 3 else "0"
                stats["total_3rd"]      = zero_if_dash(cell_texts[4]) if len(cell_texts) > 4 else "0"
                stats["total_earnings"] = cell_texts[5] if len(cell_texts) > 5 else ""
    except Exception:
        pass

    # Full race text (cleaned) → text CSV
    try:
        full_text = await accordion.inner_text()

        updated_match = re.search(r"Race Record Last Updated:\s*(.+)", full_text)
        if updated_match:
            stats["race_record_updated"] = updated_match.group(1).strip()

        # Count barrier trials before stripping markers
        stats["barrier_trial_count"] = len(
            re.findall(r"--trial--", full_text, re.IGNORECASE)
        )

        # Clean: replace --trial-- prefix with readable label, tidy whitespace
        cleaned = re.sub(r"--trial--", "[TRIAL] ", full_text, flags=re.IGNORECASE)
        text["race_record_full"] = clean_whitespace(cleaned)

    except Exception:
        stats["barrier_trial_count"] = 0
        text["race_record_full"] = ""

    return stats, text


async def extract_reports(page):
    stats = {}
    await open_accordion(page, "accordion_reports", wait_selector="p, a")
    accordion = page.locator("#accordion_reports")

    try:
        full_text = await accordion.inner_text()
        pdf_names = re.findall(r"[\w\s\-]+\.pdf", full_text, re.IGNORECASE)
        pdf_names = [n.strip() for n in pdf_names if n.strip()]
        stats["reports"]      = " | ".join(pdf_names)
        stats["report_count"] = len(pdf_names)
    except Exception:
        stats["reports"]      = ""
        stats["report_count"] = 0

    return stats


# ─────────────────────────────────────────────
# CORE LOT SCRAPER
# ─────────────────────────────────────────────

async def scrape_lot(page, sale_slug, lot_number, horse_cache=None, cache_lock=None, manifest_row=None):
    lot_id  = make_lot_id(sale_slug, lot_number)
    url     = f"{BASE_URL}/catalogue/auction/{sale_slug}/lot/{lot_number}"
    scraped = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    stats_row = {
        "lot_id":     lot_id,
        "sale_slug":  sale_slug,
        "lot":        lot_number,
        "url":        url,
        "scraped_at": scraped,
        "error":      "",
    }
    text_row = {
        "lot_id":    lot_id,
        "sale_slug": sale_slug,
        "lot":       lot_number,
    }

    print(f"  → {url}")

    try:
        await page.goto(url, timeout=30000)
        await page.wait_for_load_state("networkidle", timeout=20000)
    except Exception as e:
        stats_row["error"] = f"Page load failed: {e}"
        return stats_row, text_row

    page_text = await page.inner_text("body")
    if "not found" in page_text.lower() or "404" in page_text.lower():
        stats_row["error"] = "Lot not found (404)"
        return stats_row, text_row

    # --- Header fields (visible without accordion) ---
    try:
        lot_header = await page.locator("text=/Lot \\d+:/i").first.inner_text()
        m = re.search(r"Lot \d+:\s*(.+)", lot_header, re.IGNORECASE)
        stats_row["horse_name"] = m.group(1).strip() if m else ""
    except Exception:
        stats_row["horse_name"] = ""

    # Share % and base name — used for caching and RS category enrichment
    share_match = re.search(r"\((\d+%)\)", stats_row.get("horse_name", ""))
    share_pct   = share_match.group(1) if share_match else ""
    base_name   = re.sub(r"\s*\(\d+%\)\s*$", "", stats_row.get("horse_name", "")).strip().lower()

    # Return cached data for duplicate share lots (same horse, different %)
    if horse_cache is not None and cache_lock is not None and base_name:
        async with cache_lock:
            if base_name in horse_cache:
                cached_s, cached_t = horse_cache[base_name]
                s_row = {**cached_s,
                         "lot_id":     lot_id,
                         "lot":        lot_number,
                         "url":        url,
                         "scraped_at": scraped,
                         "horse_name": stats_row["horse_name"],
                         "category":   f"RS {share_pct}" if share_pct else cached_s.get("category", ""),
                         "error":      ""}
                t_row = {**cached_t, "lot_id": lot_id, "lot": lot_number}
                return s_row, t_row

    # Parse header "4YO G (RH)" → separate age / sex / category columns
    try:
        header_text = (await page.locator("text=/\\dYO/i").first.inner_text()).strip()
        m = re.search(r"(\d+YO)\s+([A-Z])\s+\(([A-Z]+)\)", header_text, re.IGNORECASE)
        if m:
            stats_row["age"]      = m.group(1).upper()
            stats_row["sex"]      = m.group(2).upper()
            stats_row["category"] = m.group(3).upper()
    except Exception:
        pass

    try:
        el = page.locator("p.text-gray-500").filter(
            has_text=re.compile(r"\(NZ\)|\(GB\)|\(AUS\)|\(IRE\)|\(USA\)")
        ).first
        sire_dam_raw = (await el.inner_text()).strip()
        stats_row["sire_dam_header"] = sire_dam_raw
    except Exception:
        sire_dam_raw = ""
        stats_row["sire_dam_header"] = ""

    try:
        el = page.locator(
            "span.bg-blue-200, span.bg-green-200, span.bg-red-200"
        ).first
        stats_row["reserve_status"] = (await el.inner_text()).strip()
    except Exception:
        stats_row["reserve_status"] = ""

    # --- Accordions ---
    ov_stats, ov_text = await extract_overview(page)
    pd_stats, pd_text = await extract_pedigree(page)
    rr_stats, rr_text = await extract_race_record(page)
    rp_stats          = await extract_reports(page)

    stats_row.update(ov_stats)
    stats_row.update(pd_stats)
    stats_row.update(rr_stats)
    stats_row.update(rp_stats)

    text_row.update(ov_text)
    text_row.update(pd_text)
    text_row.update(rr_text)

    # Fallback: parse sire/dam from header string if Overview table missed them
    if not stats_row.get("sire") and sire_dam_raw:
        parts = sire_dam_raw.split("/")
        if len(parts) == 2:
            stats_row["sire"] = parts[0].strip()
            stats_row["dam"]  = parts[1].strip()

    # Backfill from manifest — only fills fields the scraper left empty
    if manifest_row:
        for k, v in manifest_row.items():
            if k != "lot" and v and not stats_row.get(k):
                stats_row[k] = v

    # Enrich RS category with share percentage (e.g. "RS" → "RS 20%")
    if share_pct and stats_row.get("category", "").upper() == "RS":
        stats_row["category"] = f"RS {share_pct}"

    # Cache this horse so duplicate share lots can reuse the data
    if horse_cache is not None and cache_lock is not None and base_name:
        async with cache_lock:
            if base_name not in horse_cache:
                horse_cache[base_name] = (stats_row.copy(), text_row.copy())

    return stats_row, text_row




# ─────────────────────────────────────────────
# STARTUP — CLI args take precedence, prompts fall back when omitted.
# ─────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="Inglis Digital catalogue scraper. Run with --sale to skip prompts."
    )
    p.add_argument("--sale", help="Sale code, e.g. 2026-05B. Skips interactive prompts.")
    p.add_argument("--start", type=int, help="Start lot number (default: first in manifest)")
    p.add_argument("--end",   type=int, help="End lot number (default: last in manifest)")
    p.add_argument("--manifest", help="Override manifest path (default: {sale}_manifest.csv)")
    return p.parse_args()


def resolve_sale_inputs(args):
    """
    Return (sale_slug, manifest_dict, lot_start, lot_end).
    Uses CLI args if --sale was passed; otherwise drops into interactive prompts.
    """
    if args.sale:
        sale_slug = short_to_slug(args.sale)
        m_path    = args.manifest or manifest_path(args.sale)
        try:
            manifest = read_manifest(m_path)
        except FileNotFoundError:
            sys.exit(f"Manifest file not found: {m_path}")
        if not manifest:
            sys.exit(f"No lots in {m_path}")
        all_lots = sorted(manifest.keys())
        lot_start = args.start if args.start else all_lots[0]
        lot_end   = args.end   if args.end   else all_lots[-1]
        return sale_slug, manifest, lot_start, lot_end

    # Interactive fallback
    print("\n╔══════════════════════════════════════════╗")
    print("║   Racing Insights — Inglis Scraper       ║")
    print("╚══════════════════════════════════════════╝\n")

    sale_slug = input("Sale slug (e.g. 2026-march-early-online-sale):\n> ").strip()
    if not sale_slug:
        sys.exit("No sale slug entered.")

    m_path = input("\nManifest CSV path (e.g. 2026-03A_manifest.csv):\n> ").strip()
    if not m_path:
        sys.exit("No manifest path entered.")

    try:
        manifest = read_manifest(m_path)
    except FileNotFoundError:
        sys.exit(f"Manifest file not found: {m_path}")
    if not manifest:
        sys.exit("No lots found in manifest.")

    all_lots = sorted(manifest.keys())
    min_lot, max_lot = all_lots[0], all_lots[-1]
    print(f"\nManifest loaded: {len(manifest)} lots (#{min_lot}–#{max_lot})")

    start_in  = input(f"\nStart lot? (Enter for {min_lot}):\n> ").strip()
    lot_start = int(start_in) if start_in.isdigit() and int(start_in) > 0 else min_lot
    end_in    = input(f"End lot?   (Enter for {max_lot}):\n> ").strip()
    lot_end   = int(end_in) if end_in.isdigit() and int(end_in) > 0 else max_lot

    return sale_slug, manifest, lot_start, lot_end


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

async def run_scraper():
    args = parse_args()
    sale_slug, manifest, lot_start, lot_end = resolve_sale_inputs(args)

    short_name = slug_to_short(sale_slug)
    stats_file = stats_path(short_name)
    text_file  = text_path(short_name)

    # Only process lot numbers present in manifest within the requested range
    lot_nums    = sorted(n for n in manifest if lot_start <= n <= lot_end)
    stub_nums   = [n for n in lot_nums if manifest[n].get("category", "").upper() not in SCRAPE_CATEGORIES]
    scrape_nums = [n for n in lot_nums if n not in set(stub_nums)]
    total       = len(lot_nums)

    print(f"\nScraping '{sale_slug}' — lots {lot_start} to {lot_end}")
    print(f"  {len(scrape_nums)} to scrape  |  {len(stub_nums)} stub-only (non-RH/RF/2YO)")
    print(f"Outputs: {stats_file}  +  {text_file}\n")

    counters    = {"done": 0, "success": 0, "failed": 0, "stub": 0}
    horse_cache = {}
    cache_lock  = asyncio.Lock()
    scraped_at  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with (open(stats_file, "w", newline="", encoding="utf-8") as sf,
          open(text_file,  "w", newline="", encoding="utf-8") as tf):

        stats_w = csv.DictWriter(sf, fieldnames=STATS_FIELDS, extrasaction="ignore")
        text_w  = csv.DictWriter(tf, fieldnames=TEXT_FIELDS,  extrasaction="ignore")
        stats_w.writeheader()
        text_w.writeheader()

        # ── Write stub rows (Y/U/W) immediately — no page load needed ──
        for lot_num in stub_nums:
            m       = manifest[lot_num]
            lot_id  = make_lot_id(sale_slug, lot_num)
            stub_s  = {
                "lot_id":     lot_id,
                "sale_slug":  sale_slug,
                "lot":        lot_num,
                "url":        f"{BASE_URL}/catalogue/auction/{sale_slug}/lot/{lot_num}",
                "scraped_at": scraped_at,
                **{k: v for k, v in m.items() if k != "lot"},
                "error":      "",
            }
            stub_t = {"lot_id": lot_id, "sale_slug": sale_slug, "lot": lot_num}
            stats_w.writerow({k: stub_s.get(k, "") for k in STATS_FIELDS})
            text_w.writerow({k: stub_t.get(k, "") for k in TEXT_FIELDS})
            counters["done"] += 1
            counters["stub"] += 1
            print(f"  [{counters['done']}/{total}] Lot {lot_num}: stub ({m.get('category','?')} — {m.get('horse_name','?')})")

        sf.flush()
        tf.flush()

        write_lock = asyncio.Lock()

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 900},
            )

            page_q = asyncio.Queue()
            for _ in range(4):
                page_q.put_nowait(await context.new_page())

            async def process_lot(lot_num, idx):
                await asyncio.sleep(idx * 0.3)  # stagger initial requests
                page = await page_q.get()
                try:
                    s_row, t_row = await scrape_lot(
                        page, sale_slug, lot_num,
                        horse_cache, cache_lock,
                        manifest.get(lot_num),
                    )
                except Exception as e:
                    lot_id = make_lot_id(sale_slug, lot_num)
                    s_row = {"lot_id": lot_id, "sale_slug": sale_slug, "lot": lot_num, "error": str(e)}
                    t_row = {"lot_id": lot_id, "sale_slug": sale_slug, "lot": lot_num}
                finally:
                    page_q.put_nowait(page)

                async with write_lock:
                    counters["done"] += 1
                    err = s_row.get("error", "")
                    stats_w.writerow({k: s_row.get(k, "") for k in STATS_FIELDS})
                    text_w.writerow({k: t_row.get(k, "") for k in TEXT_FIELDS})
                    sf.flush()
                    tf.flush()
                    counters["failed" if err else "success"] += 1
                    status = err or f"✓ {s_row.get('horse_name', '—')}"
                    print(f"  [{counters['done']}/{total}] Lot {lot_num}: {status}")
                    if err:
                        log_failure(
                            short_name, "inglis_lot_scrape",
                            lot=lot_num,
                            horse_name=s_row.get("horse_name", ""),
                            status="lot_scrape_error",
                            fail_reason=err,
                        )

            await asyncio.gather(
                *[process_lot(n, i) for i, n in enumerate(scrape_nums)]
            )
            await browser.close()

    print(f"\n{'─' * 48}")
    print(f"  ✅  {counters['success']} lots scraped successfully")
    if counters["stub"]:
        print(f"  📋  {counters['stub']} lots stub-only (non-RH/RF/2YO)")
    if counters["failed"]:
        print(f"  ⚠   {counters['failed']} lots failed  (check 'error' column)")
    print(f"\n  📊  {stats_file}")
    print(f"  📄  {text_file}")
    print(f"{'─' * 48}")
    print("\n  Both files share 'lot_id' as the join key.\n")


if __name__ == "__main__":
    asyncio.run(run_scraper())
