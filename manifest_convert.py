# -*- coding: utf-8 -*-
"""
Manifest Converter — Inglis _stats → Phase 1 RA manifests

Purpose
-------
The Inglis scraper produces a rich _stats CSV containing both lot identity
and race-record data. This converter splits that into two filtered manifests:

  1. {sale}_manifest_RA.csv       — raced horses (total_runs > 0)
                                    input for 02_RA_url_discovery.py

  2. {sale}_manifest_RA_unraced.csv — unraced horses (total_runs == 0 or null)
                                    input for 02b_RA_unraced_trials.py

Both are filtered to RACEABLE_CATEGORIES (RH, RF, 2YO). Producing both from
a single run means you can sense-check both lists before committing any
Selenium time to either.

Filters
-------
1. Category whitelist (RACEABLE_CATEGORIES below).
   Whitelist is safer than blacklist: if Inglis adds a new category code,
   we'll drop it by default until you explicitly add it here.

2. total_runs > 0  → raced manifest (ratings pipeline)
   total_runs == 0 → unraced manifest (trials/trainer pipeline)

Usage
-----
    python manifest_convert.py --sale 2026-05B

Inputs
------
    {sale}_stats.csv          (output of inglisdigital_lotscraper.py)

Outputs
-------
    {sale}_manifest_RA.csv          (input for 02_RA_url_discovery.py)
    {sale}_manifest_RA_unraced.csv  (input for 02b_RA_unraced_trials.py)

The output column shape mirrors the original Inglis input manifest so both
files remain compatible with anything else built around that shape.
"""

import argparse
import csv
import sys
from collections import Counter

from sale_codes import stats_path, manifest_RA_path, manifest_RA_unraced_path


# ── Filter config ────────────────────────────────────────────────────
# Anything NOT in this set will be dropped from both outputs.
#
# Full set of category codes seen on the Inglis catalogue (May Late 2026):
#   2YO  — 2-year-olds                                  KEPT
#   RH   — Race Horse (colts/geldings/entires)          KEPT
#   RF   — Race Fillies/Mares                           KEPT
#   RS   — Race Share (fractional ownership)            dropped — same horse listed multiple
#                                                       times at different share %; lookup
#                                                       would duplicate
#   B    — Broodmare                                    dropped — racing days behind it
#   Y    — Yearling                                     dropped — not in scope
#   W    — Weanling                                     dropped — not in scope
#   SB   — Stallion-related (broodmare share?)          dropped — unclear, no race record
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


def in_scope_category(row: dict) -> bool:
    """True if the row's category is in the raceable whitelist."""
    cat = (row.get("category") or "").strip().upper()
    return cat in RACEABLE_CATEGORIES


# Category codes seen across the 13-month back catalogue (Jul 2025 - Aug 2026).
# Anything NOT in this set is genuinely new and gets called out loudly — a
# silent drop is how a whole class of lot disappears from the guide without
# anyone noticing. NOM, SB, S and U all appeared partway through the window.
KNOWN_CATEGORIES = {
    "2YO", "RH", "RF", "RS", "B", "Y", "W", "SH", "SB", "S", "U", "NOM",
}


def is_raced(row: dict) -> bool:
    """True if the horse has at least one recorded run."""
    return parse_int(row.get("total_runs")) > 0


def write_manifest(rows: list[dict], dst: str) -> None:
    with open(dst, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MANIFEST_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({
                col: (row.get(src_key, "") or "").strip()
                for col, src_key in STATS_TO_MANIFEST.items()
            })


def convert(sale: str) -> None:
    src             = stats_path(sale)
    dst_raced       = manifest_RA_path(sale)
    dst_unraced     = manifest_RA_unraced_path(sale)

    try:
        with open(src, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        sys.exit(
            f"❌ Stats file not found: {src}\n"
            f"   Run inglisdigital_lotscraper.py first."
        )

    in_scope        = [r for r in rows if in_scope_category(r)]
    out_of_scope    = len(rows) - len(in_scope)

    raced           = [r for r in in_scope if is_raced(r)]
    unraced         = [r for r in in_scope if not is_raced(r)]

    write_manifest(raced,   dst_raced)
    write_manifest(unraced, dst_unraced)

    # ── Category accounting ──────────────────────────────────────────
    # The whitelist is deliberately conservative, but a silent drop means a
    # new category could vanish from the guide unnoticed. Report what was
    # dropped, and separate "known and intentionally excluded" from "never
    # seen before, go and look at it".
    dropped_tally = Counter(
        (r.get("category") or "").strip().upper() or "(blank)"
        for r in rows if not in_scope_category(r)
    )
    unknown = {c: n for c, n in dropped_tally.items()
               if c not in KNOWN_CATEGORIES and c != "(blank)"}

    print(f"\n📂 {src}")
    print(f"   {len(rows)} total rows")
    print(f"   {out_of_scope} dropped (out-of-scope category)")
    if dropped_tally:
        detail = ", ".join(f"{c}={n}" for c, n in dropped_tally.most_common())
        print(f"     dropped breakdown: {detail}")
    print(f"   {len(in_scope)} in-scope (RH/RF/2YO)")
    print(f"     → {len(raced):>3} raced   written to {dst_raced}")
    print(f"     → {len(unraced):>3} unraced written to {dst_unraced}")

    if unknown:
        print(f"\n   {'!' * 60}")
        print(f"   !! UNRECOGNISED CATEGORY CODE(S): "
              f"{', '.join(f'{c} ({n} lots)' for c, n in unknown.items())}")
        print(f"   !! These were dropped by the whitelist. If any of them are")
        print(f"   !! raceable, add them to RACEABLE_CATEGORIES above and re-run.")
        print(f"   !! Either way add them to KNOWN_CATEGORIES so this stops firing.")
        print(f"   {'!' * 60}")

    print(f"\n   Review both files before running the RA scrape scripts.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Filter Inglis _stats CSV → raced + unraced RA manifests."
    )
    parser.add_argument("--sale", required=True, help="Sale code, e.g. 2026-05B")
    args = parser.parse_args()
    convert(args.sale)
