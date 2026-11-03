# -*- coding: utf-8 -*-
"""
merge_and_split.py
==================
Replaces the entire Power Query step in the Racing Insights pipeline.

Joins the three post-scrape CSV inputs, applies all column transformations,
computes Days Since, and writes a four-sheet Excel that build_buyers_guide_v6.py
consumes directly. No Excel, no manual steps.

Usage
-----
    python merge_and_split.py --sale 2026-05B --date 2026-06-06

    --sale   Sale code matching the other pipeline scripts, e.g. 2026-05B
    --date   Days Since baseline in YYYY-MM-DD. Pass the sale close date.
             Controls LS_days_since (maidens/winners) and Trial_Days_Since
             (unraced) — all Days Since logic lives here, not in the build script.

Inputs (all in same directory as this script)
-----
    {sale}_stats.csv              Inglis Digital scrape — all lots, all categories
                                  Output of inglisdigital_lotscraper_RH.py

    {sale}_RA_ratings.csv         Raced horse ratings + trainer
                                  Output of 03_RA_ratings_trainer.py
                                  dam_match gate already applied upstream: non-'auto'
                                  rows have blank enrichment fields in this file.

    {sale}_RA_unraced_trials.csv  Unraced horse trial + trainer
                                  Output of 02b_RA_unraced_trials.py
                                  dam_match gate applied here (see below).
                                  Optional — if absent, UNRACED trial columns are blank.

Output
------
    {sale}_merged.xlsx  Four sheets consumed by build_buyers_guide_v6.py:

    INGLIS_STATS_ALL
        Full Inglis scrape, all categories, all original columns. No RA data.
        This is the base data layer — the build script reads every lot from here.
        Kept RA-free to prevent column name collisions when the build script
        merges in MAIDENS/WINNERS/UNRACED enrichment by lot.

    MAIDENS
        RH/RF/2YO lots with 0 wins. Contains lot + horse_name + LAST_START_COLS.
        Build script merges these columns onto INGLIS_STATS_ALL rows by lot.
        horse_name is included for human readability; build script ignores it.

    WINNERS
        Same structure as MAIDENS, for lots with 1+ wins.

    UNRACED
        RH/RF/2YO lots with no race record (total_runs null or 0).
        Contains lot + horse_name + UNRACED_TRIAL_COLS.
        Trial enrichment comes from {sale}_RA_unraced_trials.csv.
        Build script merges UNRACED_TRIAL_COLS onto INGLIS_STATS_ALL rows by lot.
        Parallel pattern to how MAIDENS/WINNERS carry last-start data.

dam_match gate (unraced)
------------------------
Only rows where dam_match == 'auto' in the trials CSV have their trial
enrichment (Trainer, Stable, Last_Trial, etc.) passed through. Rows with
dam_match == 'mismatch' or 'none' get blank trial columns. Count is
reported to console. This mirrors the gate already applied by
03_RA_ratings_trainer.py for raced horses.

Track split
-----------
The RA ratings 'Track' field contains venue + going + numeric, e.g.
'RANDWICK Good 4' or 'GOULBURN Soft 5'. It is split on the LAST space
into Venue ('RANDWICK Good') and Track_rating ('4'). This replicates the
Power Query Splitter.SplitTextByEachDelimiter startAtEnd=true behaviour.

Note to future maintainers
--------------------------
Consider adding merged_path(sale) to sale_codes.py to make the output path
discoverable from other scripts, e.g.:
    merged_path(sale) -> f"{sale}_merged.xlsx"
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from sale_codes import ra_ratings_path, ra_unraced_trials_path, stats_path


# ── Column contracts (must stay in sync with build_buyers_guide_v6.py) ──────

# Columns the build script reads from MAIDENS and WINNERS sheets.
# Any change here must be mirrored in LAST_START_COLS in the build script.
LAST_START_COLS = [
    "lot",
    "Trainer",
    "Location",
    "LS_position",
    "LS_date",
    "LS_days_since",
    "Venue",
    "Track_rating",
    "Distance",
    "Class",
    "BM_Rating",
    "Rating_Juristiction",    # preserving the typo from Power Query / v5
    "BM_Rating_CountryNSW",
]

# Columns the build script reads from the UNRACED sheet.
# Any change here must be mirrored in UNRACED_TRIAL_COLS in the build script.
UNRACED_TRIAL_COLS = [
    "lot",
    "Trainer",
    "Stable",
    "Last_Trial",
    "Trial_Type",
    "Trial_Track",
    "Trial_Date",
    "Trial_Days_Since",
]

# Categories in scope — mirrors CATEGORY_FILTER in the build script
CATEGORY_FILTER = {"RH", "RF", "2YO"}

# Rename map: RA ratings column name → final name used in MAIDENS/WINNERS sheets
# 'Class' and 'Track' are not renamed (Class stays as-is; Track is split, not renamed).
RA_RENAME = {
    "trainer_name":       "Trainer",
    "trainer_location":   "Location",
    "Pos":                "LS_position",
    "Date":               "LS_date",
    "Dist":               "Distance",
    "Rating":             "BM_Rating",
    "Rating_NSW_Country": "BM_Rating_CountryNSW",
    "Rating Type":        "Rating_Juristiction",
}

# Rename map: trials CSV column name → final name used in UNRACED sheet
TRIALS_RENAME = {
    "trainer_name":     "Trainer",
    "trainer_location": "Stable",
    "last_run_result":  "Last_Trial",
    "last_run_type":    "Trial_Type",
    "last_run_track":   "Trial_Track",
    "last_run_date":    "Trial_Date",
}


# ── Loaders ──────────────────────────────────────────────────────────────────

def load_stats(path: Path) -> pd.DataFrame:
    """
    Load {sale}_stats.csv — the Inglis Digital scrape output.

    total_runs / total_wins are coerced to numeric so that blank values
    (unraced horses whose race record accordion was empty) become NaN.
    NaN is what classify_stage in the build script expects for unraced lots.
    """
    df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
    df["lot"]         = pd.to_numeric(df["lot"],         errors="coerce")
    df["total_runs"]  = pd.to_numeric(df["total_runs"],  errors="coerce")
    df["total_wins"]  = pd.to_numeric(df["total_wins"],  errors="coerce")
    df["total_2nd"]   = pd.to_numeric(df["total_2nd"],   errors="coerce")
    df["total_3rd"]   = pd.to_numeric(df["total_3rd"],   errors="coerce")
    return df


def load_ra_ratings(path: Path) -> pd.DataFrame:
    """
    Load {sale}_RA_ratings.csv — output of 03_RA_ratings_trainer.py.

    'Lot' (capital) is normalised to 'lot' (lowercase) to match the stats join key.
    'Date' is parsed as datetime for LS_days_since computation.

    The dam_match gate was already applied upstream by 03_RA_ratings_trainer.py:
    rows where dam_match != 'auto' were written with blank enrichment columns.
    No re-filtering needed here.

    Deduplicated by lot (keep last) as a safety measure against resume-logic
    edge cases that could write a lot twice.
    """
    df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
    df.rename(columns={"Lot": "lot"}, inplace=True)
    df["lot"]  = pd.to_numeric(df["lot"], errors="coerce")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Defensive \xa0 strip — automated output shouldn't carry these but guards
    # against any manual edits to the CSV between scrape and merge
    for col in ("trainer_name", "trainer_location"):
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace("\xa0", "", regex=False)
                .str.strip()
                .replace("nan", "")
            )

    df = df.drop_duplicates(subset=["lot"], keep="last")
    return df


def load_unraced_trials(
    path: Path, baseline_date: pd.Timestamp, sale: str
) -> pd.DataFrame:
    """
    Load {sale}_RA_unraced_trials.csv — output of 02b_RA_unraced_trials.py.

    Returns an empty UNRACED_TRIAL_COLS-shaped DataFrame if the file is absent
    (i.e. 02b hasn't been run yet). Unraced tab will still be written but with
    blank trial columns — the build completes, just without trial enrichment.

    dam_match gate
    --------------
    Only rows where dam_match == 'auto' have their trial enrichment passed through.
    Rows with dam_match == 'mismatch' or 'none' get blank trial columns in the output.
    A console count is printed for visibility. This mirrors the behaviour of
    03_RA_ratings_trainer.py for raced horses.

    Trial_Days_Since
    ----------------
    Computed from last_run_date (DD/MM/YYYY string) and the --date baseline.
    Null where last_run_date is blank (horse has no form on RA — expected).
    """
    empty = pd.DataFrame(columns=UNRACED_TRIAL_COLS)

    if not path.exists():
        print(
            f"  ⚠️  {path.name} not found — UNRACED trial columns will be blank.\n"
            f"     Run: python 02b_RA_unraced_trials.py --sale {sale}"
        )
        return empty

    df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
    if df.empty:
        return empty

    df.rename(columns={"Lot": "lot"}, inplace=True)
    df["lot"] = pd.to_numeric(df["lot"], errors="coerce")

    # Defensive \xa0 strip (same rationale as load_ra_ratings above)
    for col in ("trainer_name", "trainer_location"):
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace("\xa0", "", regex=False)
                .str.strip()
                .replace("nan", "")
            )

    df = df.drop_duplicates(subset=["lot"], keep="last")

    # dam_match gate: blank trial fields for non-auto rows before renaming
    non_auto_mask = df["dam_match"].astype(str).str.strip().str.lower() != "auto"
    non_auto_count = int(non_auto_mask.sum())
    if non_auto_count:
        print(
            f"  ℹ️  {non_auto_count} unraced lot(s) excluded from trial enrichment "
            f"(dam_match != 'auto') — review {path.name}"
        )
    cols_to_blank = [c for c in TRIALS_RENAME if c in df.columns]
    df.loc[non_auto_mask, cols_to_blank] = ""

    # Rename to canonical output column names
    df.rename(columns=TRIALS_RENAME, inplace=True)

    # Compute Trial_Days_Since — kept here (not in build script) so that the
    # --date baseline is configured in exactly one place for all Days Since logic
    if "Trial_Date" in df.columns:
        parsed_dates = pd.to_datetime(df["Trial_Date"], format="%d/%m/%Y", errors="coerce")
        df["Trial_Days_Since"] = (baseline_date - parsed_dates).dt.days.astype("Int64")
        # Null out where date was blank (no form) or the row was non-auto
        df.loc[parsed_dates.isna(), "Trial_Days_Since"] = pd.NA
        df.loc[non_auto_mask, "Trial_Days_Since"] = pd.NA
    else:
        df["Trial_Days_Since"] = pd.NA

    # Return only the columns the build script expects (adds NaN for any missing)
    for col in UNRACED_TRIAL_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    return df[UNRACED_TRIAL_COLS]


# ── Join and transform ────────────────────────────────────────────────────────

def build_combined(
    stats: pd.DataFrame,
    ra_ratings: pd.DataFrame,
    baseline_date: pd.Timestamp,
) -> pd.DataFrame:
    """
    Left-join stats onto RA ratings, rename enrichment columns, compute
    LS_days_since, and split the Track field.

    Unraced horses have no row in the ratings CSV so all RA columns are NaN
    for them — that is correct and expected.

    Track split
    -----------
    RA 'Track' format: 'VENUE_CODE Going Numeric' e.g. 'RANDWICK Good 4'.
    Split on last space → Venue = 'RANDWICK Good', Track_rating = '4'.
    Uses rsplit(n=1) which splits at the rightmost space — replicating
    Power Query's Splitter.SplitTextByEachDelimiter startAtEnd=true.
    Where Track is null (unraced or no last-start data), both columns are null.
    """
    # Select only the RA columns we need for the join; avoid carrying through
    # the full wide ratings CSV (lot_id, sale_slug, etc. are not needed)
    ra_cols_needed = (
        ["lot", "Track", "Class"]
        + list(RA_RENAME.keys())
    )
    ra_subset = ra_ratings[[c for c in ra_cols_needed if c in ra_ratings.columns]].copy()

    combined = stats.merge(ra_subset, on="lot", how="left")

    # Rename RA enrichment columns
    combined.rename(columns=RA_RENAME, inplace=True)

    # Compute LS_days_since — Days Since is always calculated here from --date,
    # not hardcoded per-sale as it was in Power Query (#date(2026, 5, 22)).
    if "LS_date" in combined.columns:
        combined["LS_days_since"] = (
            (baseline_date - combined["LS_date"]).dt.days.astype("Int64")
        )
        combined.loc[combined["LS_date"].isna(), "LS_days_since"] = pd.NA
    else:
        combined["LS_days_since"] = pd.NA

    # Track split: rsplit on last space replicates PQ startAtEnd=true
    track_col = combined["Track"] if "Track" in combined.columns else pd.Series(
        "", index=combined.index, dtype=str
    )
    track_split = track_col.fillna("").str.rsplit(" ", n=1, expand=True)
    # expand=True with n=1 always yields at least column 0; column 1 only appears
    # when at least one value contains a space — guard for that edge case
    if 1 not in track_split.columns:
        track_split[1] = pd.NA

    combined["Venue"]        = track_split[0].str.strip()
    combined["Track_rating"] = track_split[1].astype(str).str.strip()

    # Null out where original Track was absent
    no_track = track_col.isna() | (track_col.astype(str).str.strip() == "")
    combined.loc[no_track, ["Venue", "Track_rating"]] = pd.NA
    combined["Venue"]        = combined["Venue"].replace("", pd.NA)
    combined["Track_rating"] = combined["Track_rating"].replace({"": pd.NA, "nan": pd.NA})

    # Ensure every column that LAST_START_COLS requires is present — guards
    # against an unexpected gap in the ratings CSV producing a KeyError downstream
    for col in LAST_START_COLS:
        if col not in combined.columns:
            combined[col] = pd.NA

    return combined


# ── Stage masks ───────────────────────────────────────────────────────────────

def unraced_mask(df: pd.DataFrame) -> pd.Series:
    """total_runs is NaN or 0 — mirrors classify_stage in build_buyers_guide_v6.py."""
    return df["total_runs"].isna() | (df["total_runs"] == 0)


def winner_mask(df: pd.DataFrame) -> pd.Series:
    """total_wins is not null and > 0."""
    return df["total_wins"].notna() & (df["total_wins"] > 0)


# ── Main ──────────────────────────────────────────────────────────────────────

def main(sale: str, date_str: str) -> None:
    baseline_date = pd.Timestamp(date_str)

    stats_file   = Path(stats_path(sale))
    ratings_file = Path(ra_ratings_path(sale))
    trials_file  = Path(ra_unraced_trials_path(sale))
    output_file  = stats_file.parent / f"{sale}_merged.xlsx"

    # ── Validate required inputs ──────────────────────────────────────
    for f, label in [
        (stats_file,   "inglisdigital_lotscraper_RH.py"),
        (ratings_file, "03_RA_ratings_trainer.py"),
    ]:
        if not f.exists():
            sys.exit(
                f"❌ Required input not found: {f}\n"
                f"   Run {label} --sale {sale} first."
            )

    # ── Load ─────────────────────────────────────────────────────────
    print(f"Reading: {stats_file.name}")
    stats = load_stats(stats_file)

    print(f"Reading: {ratings_file.name}")
    ra_ratings = load_ra_ratings(ratings_file)

    print(f"Reading: {trials_file.name}")
    unraced_trials = load_unraced_trials(trials_file, baseline_date, sale)

    # ── Join and transform ────────────────────────────────────────────
    print("\nJoining and transforming...")
    combined = build_combined(stats, ra_ratings, baseline_date)

    # ── Stage classification (category-scoped) ────────────────────────
    in_scope = combined["category"].isin(CATEGORY_FILTER)
    scoped   = combined[in_scope].copy()

    is_unraced = unraced_mask(scoped)
    is_winner  = winner_mask(scoped)
    is_maiden  = ~is_unraced & ~is_winner

    # ── Build INGLIS_STATS_ALL ────────────────────────────────────────
    # Pure Inglis scrape — all categories, all original stats columns, no RA data.
    # Kept RA-free so the build script's merge-by-lot doesn't produce duplicate
    # column names (e.g. a 'Trainer' in RAW would collide with the 'Trainer'
    # merged in from MAIDENS/WINNERS).
    inglis_stats_all = stats.sort_values("lot").reset_index(drop=True)

    # ── Build MAIDENS and WINNERS ─────────────────────────────────────
    # These sheets carry only the join columns the build script reads.
    # horse_name is included for human readability when inspecting the file;
    # the build script ignores it (it selects LAST_START_COLS explicitly).
    last_start_sheet_cols = ["horse_name"] + LAST_START_COLS

    maidens_df = (
        scoped[is_maiden][last_start_sheet_cols]
        .sort_values("lot")
        .reset_index(drop=True)
    )
    winners_df = (
        scoped[is_winner][last_start_sheet_cols]
        .sort_values("lot")
        .reset_index(drop=True)
    )

    # ── Build UNRACED ─────────────────────────────────────────────────
    # Base identity columns from stats + trial enrichment from 02b.
    # Build script merges UNRACED_TRIAL_COLS onto INGLIS_STATS_ALL rows by lot —
    # the same join pattern used for MAIDENS/WINNERS last-start data.
    unraced_base = (
        scoped[is_unraced][["lot", "horse_name"]]
        .sort_values("lot")
        .reset_index(drop=True)
    )

    if not unraced_trials.empty:
        unraced_df = unraced_base.merge(unraced_trials, on="lot", how="left")
    else:
        unraced_df = unraced_base.copy()

    # Ensure every UNRACED_TRIAL_COLS column is present even if trials were absent
    for col in UNRACED_TRIAL_COLS:
        if col not in unraced_df.columns:
            unraced_df[col] = pd.NA

    # ── Write output ──────────────────────────────────────────────────
    print(f"\nWriting: {output_file.name}")
    with pd.ExcelWriter(output_file, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
        inglis_stats_all.to_excel(writer, sheet_name="INGLIS_STATS_ALL", index=False)
        maidens_df.to_excel(writer, sheet_name="MAIDENS",          index=False)
        winners_df.to_excel(writer, sheet_name="WINNERS",          index=False)
        unraced_df.to_excel(writer, sheet_name="UNRACED",          index=False)

    # ── Summary ───────────────────────────────────────────────────────
    n_unraced = int(is_unraced.sum())
    n_maiden  = int(is_maiden.sum())
    n_winner  = int(is_winner.sum())
    n_trial_enriched = int(
    (unraced_df["Trainer"].notna() & (unraced_df["Trainer"].astype(str).str.strip() != ""))
    .sum()
) if "Trainer" in unraced_df.columns else 0

    print(f"\n{'─' * 52}")
    print(f"  ✅  {output_file.name}")
    print(f"{'─' * 52}")
    print(f"  In scope (RH / RF / 2YO):  {int(in_scope.sum())} lots")
    print(f"    UNRACED  {n_unraced:>4d}   "
          f"({n_trial_enriched} with trial enrichment)")
    print(f"    MAIDENS  {n_maiden:>4d}")
    print(f"    WINNERS  {n_winner:>4d}")
    print(f"\n  Sheets written:")
    print(f"    INGLIS_STATS_ALL  {len(inglis_stats_all):>4d} rows  (all categories)")
    print(f"    MAIDENS           {len(maidens_df):>4d} rows")
    print(f"    WINNERS           {len(winners_df):>4d} rows")
    print(f"    UNRACED           {len(unraced_df):>4d} rows")
    print(f"\n  Days Since baseline: {date_str}")
    print(f"{'─' * 52}")
    print(
        f"\n  Next: python build_buyers_guide_v6.py "
        f"--sale {sale} --date {date_str}"
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Merge Inglis + RA data and split into the four-sheet Excel consumed "
            "by build_buyers_guide_v6.py. Replaces the Power Query step."
        )
    )
    parser.add_argument(
        "--sale", required=True,
        help="Sale code, e.g. 2026-05B"
    )
    parser.add_argument(
        "--date", required=True,
        help=(
            "Days Since baseline date, YYYY-MM-DD, e.g. 2026-06-06. "
            "Pass the sale close date. Controls LS_days_since (maidens/winners) "
            "and Trial_Days_Since (unraced)."
        ),
    )
    args = parser.parse_args()

    try:
        pd.Timestamp(args.date)
    except Exception:
        sys.exit(f"❌ --date must be YYYY-MM-DD, got: {args.date!r}")

    main(args.sale, args.date)
