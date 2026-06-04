# -*- coding: utf-8 -*-
"""
Manifest Converter — Inglis _stats → Phase 1 RA manifest

Purpose
-------
The Inglis scraper produces a rich _stats CSV containing both lot identity
and race-record data. Phase 1 (RA URL discovery) only needs a small subset
of those rows: horses that have actually raced and therefore could have a
Racing Australia rating to find.

This converter applies that filter, eliminating wasted Selenium searches
before Phase 1 ever starts. On a typical sale this drops 30–50% of lots.

Filters
-------
1. Category whitelist (RACEABLE_CATEGORIES below).
   Whitelist is safer than blacklist: if Inglis adds a new category code,
   we'll drop it by default until you explicitly add it here.

2. total_runs > 0.
   Unraced horses cannot have a Racing Australia rating, so searching for
   them is pure waste. Phase 2 would return empty anyway.

Usage
-----
    python manifest_convert.py --sale 2026-05B

Inputs
------
    {sale}_stats.csv          (output of inglisdigital_lotscraper.py)

Outputs
-------
    {sale}_manifest_RA.csv    (input for 02_RA_url_discovery.py)

The output column shape mirrors the original Inglis input manifest so the
filtered file remains compatible with anything else built around that shape.
"""

import argparse
import csv
import sys

from sale_codes import stats_path, manifest_RA_path


# ── Filter config ────────────────────────────────────────────────────
# Anything NOT in this set will be dropped. Edit if you want to include
# more.
#
# Full set of category codes seen on the Inglis catalogue (May Late 2026):
#   2YO  — 2-year-olds                                  KEPT
#   RH   — Race Horse (colts/geldings/entires)          KEPT
#   RF   — Race Fillies/Mares                           KEPT
#   RS   — Race Share (fractional ownership)            dropped — same horse listed multiple
#                                                       times at different share %; rating
#                                                       lookup would duplicate
#   B    — Broodmare                                    dropped — racing days behind it
#   Y    — Yearling                                     dropped — unraced
#   W    — Weanling                                     dropped — unraced
#   SB   — Stallion-related (broodmare share?)          dropped — unclear, no current race record
#   SH   — Stallion / share                             dropped — no race record sought
#
# Adding RS would also need de-duplication logic (same horse, multiple lots).
RACEABLE_CATEGORIES = {"2YO", "RH", "RF"}


# ── Output shape ─────────────────────────────────────────────────────
# Matches the column order of the original hand-built manifest, which is
# what the Inglis scraper's read_manifest() expects and what Phase 1 reads.
MANIFEST_COLUMNS = [
    "Lot", "Age", "Cat", "Sex", "Name",
    "Sire", "Dam", "Vendor", "Covering Stallion", "State",
]

# Map output column → source key in the _stats CSV.
STATS_TO_MANIFEST = {
    "Lot":               "lot",
    "Age":               "age",
    "Cat":               "category",
    "Sex":               "sex",
    "Name":              "horse_name",
    "Sire":              "sire",
    "Dam":               "dam",
    "Vendor":            "vendor",
    "Covering Stallion": "covering_stallion",
    "State":             "state",
}


def parse_int(value) -> int:
    """Inglis sometimes writes '-' for zero. Treat anything non-numeric as 0."""
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return 0


def should_keep(row: dict) -> tuple[bool, str]:
    """Apply filters; return (keep, reason_if_dropped)."""
    cat = (row.get("category") or "").strip().upper()
    if cat not in RACEABLE_CATEGORIES:
        return False, f"category={cat or 'blank'}"

    if parse_int(row.get("total_runs")) == 0:
        return False, "zero starts"

    return True, ""


def convert(sale: str) -> None:
    src = stats_path(sale)
    dst = manifest_RA_path(sale)

    try:
        with open(src, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        sys.exit(
            f"❌ Stats file not found: {src}\n"
            f"   Run inglisdigital_lotscraper.py first."
        )

    kept = []
    dropped_counts: dict[str, int] = {}
    for row in rows:
        keep, reason = should_keep(row)
        if keep:
            kept.append(row)
        else:
            dropped_counts[reason] = dropped_counts.get(reason, 0) + 1

    with open(dst, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS)
        writer.writeheader()
        for row in kept:
            writer.writerow({
                col: (row.get(src_key, "") or "").strip()
                for col, src_key in STATS_TO_MANIFEST.items()
            })

    print(f"\n📂 {src} → {dst}")
    print(f"   {len(rows)} stats rows → {len(kept)} kept for RA lookup")
    if dropped_counts:
        print("   Dropped:")
        for reason, count in sorted(dropped_counts.items(), key=lambda x: -x[1]):
            print(f"     {count:>4}  {reason}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Filter Inglis _stats CSV to raceable horses, emit RA manifest."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    args = parser.parse_args()
    convert(args.sale)
