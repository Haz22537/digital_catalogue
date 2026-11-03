"""
build_buyers_guide_v6.py
========================
Builds the Racing Insights Buyers' Guide from the merged Excel produced by
merge_and_split.py.

v6 changes from v5
------------------
  * INPUT_FILE, OUTPUT_FILE, SALE_TITLE, CATALOGUE_REF, PREPARED_DATE are now
    derived from --sale and --date CLI arguments. The per-sale config block that
    required a manual edit before every run has been removed.

  * RAW_SHEET renamed to "INGLIS_STATS_ALL" (was "RAW_UNFILTERED") to match the
    sheet name written by merge_and_split.py.

  * UNRACED tab expanded from 11 to 18 columns. The seven new columns —
    Trainer, Stable, Last Trial, Type, TRACK, Date, Days Since — are read from
    the UNRACED sheet in the merged Excel and merged onto the base data by lot.
    This is the same join pattern used for last-start data in MAIDENS/WINNERS.

  * Days Since (Trial_Days_Since) gets the same three-band fill as LS_days_since.
    Last Trial (Last_Trial) gets the same win/place fill as LS_position.

  * All v5 branding, formatting, and visual layer is preserved unchanged.

Usage
-----
    python build_buyers_guide_v6.py --sale 2026-05B --date 2026-06-06

    --sale   Sale code, e.g. 2026-05B
    --date   Prepared date in YYYY-MM-DD, e.g. 2026-06-06 (displayed as "6 June 2026")

Input
-----
    {sale}_merged.xlsx    Output of merge_and_split.py. Must contain four sheets:
                          INGLIS_STATS_ALL, MAIDENS, WINNERS, UNRACED.

Output
------
    BuyersGuide_{sale}_{date}.xlsx
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from sale_codes import MONTHS, TIMING_WORDS


# =============================================================================
# STATIC CONFIG — does not change per sale
# =============================================================================
LOGO_FILE        = "logo_transparent_full.png"
CATEGORY_FILTER  = ["RH", "RF", "2YO"]
STRIP_AUS_SUFFIX = True

# Sheet names in the merged Excel (must match merge_and_split.py)
RAW_SHEET            = "INGLIS_STATS_ALL"   # v6: was "RAW_UNFILTERED" in v5
MAIDENS_SOURCE_SHEET = "MAIDENS"
WINNERS_SOURCE_SHEET = "WINNERS"
UNRACED_SOURCE_SHEET = "UNRACED"            # v6: new — carries trial enrichment


# =============================================================================
# TOP METRO TRAINERS — edit to curate
# =============================================================================
TOP_METRO_TRAINERS = {
    # Existing entries
    "chris waller",
    "annabel neasham",
    "ciaron maher",
    "anthony freedman",
    "sam freedman",
    "anthony & sam freedman",
    "james cummings",
    "godolphin",
    "peter & paul snowden",
    "snowden racing",
    "gai waterhouse",
    "adrian bott",
    "waterhouse & bott",
    "mick price",
    "michael kent jnr",
    "price & kent",
    "danny o'brien",
    "matthew ellerton",
    "simon zahra",
    "ellerton zahra",
    "tony & calvin mcevoy",
    "mcevoy mitchell",
    "phillip stokes",
    "lindsay park",
    "ben hayes",
    "j d hayes",
    "hayes racing",
    "michael, wayne & john hawkes",
    "hawkes racing",
    "kris lees",
    "joseph pride",
    "bjorn baker",
    
    # New entries from your list (35 additions)
    "ben, will & j d hayes",
    "gai waterhouse & adrian bott",
    "tony gollan",
    "mick price & michael kent jnr",
    "annabel & rob archibald",
    "michael freedman",
    "peter g moody & katherine coleman",
    "robert heathcote",
    "grant & alana williams",
    "chris & corey munce",
    "michael grantham",
    "richard & chantelle jolly",
    "andrew gluyas",
    "will clarken",
    "grahame begg",
    "matt laurie",
    "leon & troy corstens & will larkin",
    "gavin bedggood",
    "trent busuttin & natalie young",
    "will hayes",
    "annabel archibald",
    "peter g moody",
    "grant williams",
    "chris munce",
    "richard jolly",
    "leon corstens",
    "trent busuttin",
    "will larkin",
    "rob archibald",
    "katherine coleman",
    "alana williams",
    "corey munce",
    "chantelle jolly",
    "troy corstens",
    "natalie young",
}


# =============================================================================
# BRANDING — v5 palette (unchanged)
# =============================================================================
CHARCOAL_NAVY = "0D1117"
MINT          = "00C896"
WHITE         = "FFFFFF"
ROW_RULE_GREY = "E5E7EB"
SOFT_GREY     = "6B7280"
FONT_NAME     = "Calibri"

NAVY = CHARCOAL_NAVY

NAVY_FILL = PatternFill("solid", fgColor=CHARCOAL_NAVY)
WHITE_FILL = PatternFill("solid", fgColor=WHITE)
MINT_FILL  = PatternFill("solid", fgColor=MINT)

ROW_BORDER        = Border(bottom=Side(style="thin", color=ROW_RULE_GREY))
MINT_ACCENT_BORDER = Border(bottom=Side(style="medium", color=MINT))

TITLE_FONT    = Font(name=FONT_NAME, size=16, bold=True,  color=WHITE)
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, color=WHITE, italic=True)
HEADER_FONT   = Font(name=FONT_NAME, size=11, bold=True,  color=WHITE)
BODY_FONT     = Font(name=FONT_NAME, size=10, color=NAVY)
BODY_BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
HYPERLINK_FONT = Font(name=FONT_NAME, size=10, color=NAVY, underline="single")
FOOTER_FONT   = Font(name=FONT_NAME, size=9,  italic=True, color=NAVY)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")


# =============================================================================
# VISUAL LAYER — unchanged from v5
# =============================================================================
DS_GREEN_FILL = PatternFill("solid", fgColor="CDE9D2")
DS_AMBER_FILL = PatternFill("solid", fgColor="FCEFC9")
DS_RED_FILL   = PatternFill("solid", fgColor="F5D5D0")

LS_WIN_FILL   = PatternFill("solid", fgColor="D9EBE0")
LS_PLACE_FILL = PatternFill("solid", fgColor="EAEFE2")

RATING_GREEN_THRESHOLD = 75
RATING_GREY_THRESHOLD  = 60
RATING_GREEN_FONT = Font(name=FONT_NAME, size=10, bold=True, color="1F7A3A")
RATING_GREY_FONT  = Font(name=FONT_NAME, size=10, color="7A8290")

TRAINER_STAR      = "★ "
TRAINER_STAR_FONT = Font(name=FONT_NAME, size=10, bold=True, color=MINT)

REPORT_PRESENT = "📄"
REPORT_NONE    = ""

EARNINGS_BAR_COLOR = MINT


# =============================================================================
# COLUMN SPECS
# =============================================================================
# v6: UNRACED expanded from 11 to 18 columns.
# The 7 new columns (Trainer → Days Since) are inserted between Trials and
# Reports — matching the layout of BuyersGuide_202605B_late_May_2026_v6.xlsx.
# MAIDENS and WINNERS are unchanged from v5.
UNRACED_COLS = [
    ("Lot",         "lot",                CENTER,  6),
    ("Horse Name",  "horse_name",         LEFT,   22),
    ("Reserve",     "reserve",            CENTER, 11),
    ("Age",         "age",                CENTER,  7),
    ("Sex",         "sex",                CENTER,  6),
    ("Sire",        "sire",               LEFT,   20),
    ("Dam",         "dam",                LEFT,   20),
    ("Vendor",      "vendor",             LEFT,   24),
    ("State",       "state",              CENTER,  8),
    ("Trials",      "barrier_trial_count",CENTER,  8),
    # v6: trial enrichment columns (from UNRACED sheet in merged Excel)
    ("Trainer",     "Trainer",            LEFT,   22),
    ("Stable",      "Stable",             LEFT,   16),
    ("Last Trial",  "Last_Trial",         CENTER, 12),
    ("Type",        "Trial_Type",         CENTER,  8),
    ("TRACK",       "Trial_Track",        CENTER,  8),
    ("Date",        "Trial_Date",         CENTER, 12),
    ("Days Since",  "Trial_Days_Since",   CENTER, 11),
    ("Reports",     "reports_flag",       CENTER,  9),
]

MAIDENS_COLS = [
    ("Lot",         "lot",                CENTER,  6),
    ("Horse Name",  "horse_name",         LEFT,   22),
    ("Reserve",     "reserve",            CENTER, 11),
    ("Age",         "age",                CENTER,  7),
    ("Sex",         "sex",                CENTER,  6),
    ("Sire",        "sire",               LEFT,   20),
    ("Dam",         "dam",                LEFT,   20),
    ("Vendor",      "vendor",             LEFT,   24),
    ("State",       "state",              CENTER,  8),
    ("Starts",      "total_runs",         CENTER,  8),
    ("Record",      "record_w23",         CENTER, 10),
    ("Earnings",    "total_earnings",     RIGHT,  11),
    ("Trainer",     "Trainer",            LEFT,   20),
    ("Last Start",  "LS_position",        CENTER, 11),
    ("Days Since",  "LS_days_since",      CENTER, 11),
    ("Last Class",  "Class",              CENTER, 11),
    ("Last Dist",   "Distance",           CENTER, 10),
    ("Trials",      "barrier_trial_count",CENTER,  8),
    ("Reports",     "reports_flag",       CENTER,  9),
]

WINNERS_COLS = [
    ("Lot",          "lot",                CENTER,  6),
    ("Horse Name",   "horse_name",         LEFT,   22),
    ("Reserve",      "reserve",            CENTER, 11),
    ("Age",          "age",                CENTER,  7),
    ("Sex",          "sex",                CENTER,  6),
    ("Sire",         "sire",               LEFT,   20),
    ("Dam",          "dam",                LEFT,   20),
    ("Vendor",       "vendor",             LEFT,   24),
    ("State",        "state",              CENTER,  8),
    ("Starts",       "total_runs",         CENTER,  8),
    ("Record",       "record_w23",         CENTER, 10),
    ("Earnings",     "total_earnings",     RIGHT,  14),
    ("Trainer",      "Trainer",            LEFT,   20),
    ("Last Start",   "LS_position",        CENTER, 11),
    ("Days Since",   "LS_days_since",      CENTER, 11),
    ("Last Class",   "Class",              CENTER, 11),
    ("Last Dist",    "Distance",           CENTER, 10),
    ("Rating",       "rating_display",     CENTER,  9),
    ("Jurisdiction", "jurisdiction_display",CENTER, 14),
    ("Reports",      "reports_flag",       CENTER,  9),
]

# Columns read from the UNRACED sheet in the merged Excel.
# Must stay in sync with UNRACED_TRIAL_COLS in merge_and_split.py.
UNRACED_TRIAL_COLS = [
    "lot", "Trainer", "Stable", "Last_Trial", "Trial_Type",
    "Trial_Track", "Trial_Date", "Trial_Days_Since",
]

# Columns read from MAIDENS and WINNERS sheets in the merged Excel.
LAST_START_COLS = [
    "lot", "Trainer", "Location", "LS_position", "LS_date", "LS_days_since",
    "Venue", "Track_rating", "Distance", "Class", "BM_Rating",
    "Rating_Juristiction",    # preserving the typo from Power Query / v5
    "BM_Rating_CountryNSW",
]


# =============================================================================
# SALE-CODE HELPERS — derive display strings from the --sale and --date args
# =============================================================================
def sale_title_from_code(sale: str) -> str:
    """
    Derive a human-readable sale title from the short sale code.
    '2026-05B' → 'May (Late) Online Sale — Buyer\u2019s Guide'
    Falls back gracefully if the code is in an unexpected format.
    """
    try:
        _, monthcode = sale.split("-", 1)
        month_num = monthcode[:2]
        timing    = monthcode[2:].upper()
        month_name  = MONTHS.get(month_num, month_num).capitalize()
        timing_word = TIMING_WORDS.get(timing, timing).capitalize()
        return f"{month_name} ({timing_word}) Online Sale \u2014 Buyer\u2019s Guide"
    except Exception:
        return f"{sale} \u2014 Buyer\u2019s Guide"


def prepared_date_from_str(date_str: str) -> str:
    """
    Format a YYYY-MM-DD string as a display date.
    '2026-06-06' → '6 June 2026'
    """
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{d.day} {d.strftime('%B %Y')}"


# =============================================================================
# PREPROCESSING
# =============================================================================
def classify_stage(row):
    runs = row["total_runs"]
    wins = row["total_wins"]
    if pd.isna(runs) or runs == 0:
        return "UNRACED"
    if pd.isna(wins) or wins == 0:
        return "MAIDENS"
    return "WINNERS"


def resolve_rating(row):
    bm      = row["BM_Rating"]
    country = row["BM_Rating_CountryNSW"]
    jur     = row["Rating_Juristiction"]
    if pd.notna(bm) and bm != 0:
        return bm, shorten_jurisdiction(jur, country=False)
    if pd.notna(country) and country != 0:
        return country, shorten_jurisdiction(jur, country=True)
    return None, shorten_jurisdiction(jur, country=False)


def shorten_jurisdiction(jur, country=False):
    if pd.isna(jur):
        return "NSW (Country)" if country else None
    s = str(jur).strip()
    s = re.sub(r"\s+Benchmark$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+Rating$",    "", s, flags=re.IGNORECASE)
    return f"{s} (Country)" if country else s


def parse_finish_position(ls_str):
    if pd.isna(ls_str):
        return None
    m = re.match(r"^(\d+)(?:st|nd|rd|th)?\s+of\s+\d+", str(ls_str).strip(), re.IGNORECASE)
    return int(m.group(1)) if m else None


def is_top_trainer(trainer_name):
    if pd.isna(trainer_name):
        return False
    t = str(trainer_name).lower().strip()
    return any(top in t for top in TOP_METRO_TRAINERS)


def preprocess(input_path):
    # ── Base data ──────────────────────────────────────────────────────
    raw = pd.read_excel(input_path, sheet_name=RAW_SHEET)

    # ── Last-start enrichment for raced horses (MAIDENS / WINNERS) ────
    ls_frames = []
    for sheet in (MAIDENS_SOURCE_SHEET, WINNERS_SOURCE_SHEET):
        if sheet is None:
            continue
        try:
            ls_frames.append(
                pd.read_excel(input_path, sheet_name=sheet)[LAST_START_COLS]
            )
        except (ValueError, KeyError) as e:
            print(f"  Warning: could not load last-start data from '{sheet}': {e}")

    ls_data = (
        pd.concat(ls_frames, ignore_index=True)
        if ls_frames else pd.DataFrame(columns=LAST_START_COLS)
    )

    # ── Trial enrichment for unraced horses (UNRACED) ─────────────────
    # v6: reads Trainer, Stable, Last_Trial, Trial_Type, Trial_Track,
    # Trial_Date, Trial_Days_Since from the UNRACED sheet.
    try:
        trial_data = pd.read_excel(
            input_path, sheet_name=UNRACED_SOURCE_SHEET
        )[UNRACED_TRIAL_COLS]
    except (ValueError, KeyError) as e:
        print(f"  Warning: could not load trial data from '{UNRACED_SOURCE_SHEET}': {e}")
        trial_data = pd.DataFrame(columns=UNRACED_TRIAL_COLS)

    # ── Filter to in-scope categories ─────────────────────────────────
    df = raw[raw["category"].isin(CATEGORY_FILTER)].copy()

    # ── Merge last-start enrichment ────────────────────────────────────
    df = df.merge(ls_data, on="lot", how="left")

    # ── Merge trial enrichment ─────────────────────────────────────────
    # 'Trainer' appears in both ls_data (for raced horses) and trial_data
    # (for unraced horses). Because unraced lots are absent from the MAIDENS/
    # WINNERS sheets, their Trainer is null after the ls_data merge. The
    # fillna below populates it from trial_data without touching the already-
    # populated Trainer values on MAIDENS/WINNERS rows.
    if not trial_data.empty:
        df = df.merge(trial_data, on="lot", how="left", suffixes=("", "_unraced"))
        if "Trainer_unraced" in df.columns:
            df["Trainer"] = df["Trainer"].fillna(df["Trainer_unraced"])
            df.drop(columns=["Trainer_unraced"], inplace=True)

    # Ensure every trial column exists even if the UNRACED sheet was absent
    for col in [c for c in UNRACED_TRIAL_COLS if c != "lot"]:
        if col not in df.columns:
            df[col] = pd.NA

    # ── Standard derived columns ───────────────────────────────────────
    if STRIP_AUS_SUFFIX:
        df["horse_name"] = (
            df["horse_name"].astype(str)
            .str.replace(r"\s*\(AUS\)\s*$", "", regex=True).str.strip()
        )

    df["reports_flag"] = df["report_count"].apply(
        lambda x: REPORT_PRESENT if (pd.notna(x) and x > 0) else REPORT_NONE
    )
    df["reserve"] = df["reserve_status"].fillna("").apply(
        lambda x: "Unreserved" if x == "Unreserved" else ""
    )
    df["record_w23"] = df.apply(
        lambda r: (
            f"{int(r['total_wins']) if pd.notna(r['total_wins']) else 0}-"
            f"{int(r['total_2nd'])  if pd.notna(r['total_2nd'])  else 0}-"
            f"{int(r['total_3rd'])  if pd.notna(r['total_3rd'])  else 0}"
        ),
        axis=1,
    )

    # Resolve rating — requires BM_Rating / BM_Rating_CountryNSW / Rating_Juristiction
    # These are null for UNRACED horses; resolve_rating handles that gracefully.
    for col in ("BM_Rating", "BM_Rating_CountryNSW", "Rating_Juristiction"):
        if col not in df.columns:
            df[col] = pd.NA

    rating_pairs = df.apply(resolve_rating, axis=1)
    df["rating_display"]      = [p[0] for p in rating_pairs]
    df["jurisdiction_display"] = [p[1] for p in rating_pairs]

    df["stage"]          = df.apply(classify_stage, axis=1)
    df["is_top_trainer"] = df["Trainer"].apply(is_top_trainer)

    df["total_earnings"] = df["total_earnings"].apply(
        lambda v: "A$0" if (v == 0 or v == "0" or v == 0.0) else v
    )

    df = df.sort_values("lot").reset_index(drop=True)

    print(f"  Filtered to {len(df)} lots ({', '.join(CATEGORY_FILTER)} categories)")
    print(f"    UNRACED: {(df['stage'] == 'UNRACED').sum()}")
    print(f"    MAIDENS: {(df['stage'] == 'MAIDENS').sum()}")
    print(f"    WINNERS: {(df['stage'] == 'WINNERS').sum()}")

    # ── Action items ───────────────────────────────────────────────────
    print()
    print("=" * 70)
    print("ACTION ITEMS — please review before publishing:")
    print("=" * 70)

    w = df[df["stage"] == "WINNERS"]
    blank_ratings = w[w["rating_display"].isna()]
    if len(blank_ratings) > 0:
        print(f"\n* {len(blank_ratings)} WINNERS have blank ratings. Lots to review:")
        for _, r in blank_ratings.iterrows():
            print(
                f"    Lot {int(r['lot']):3d}  {r['horse_name']:20s}  "
                f"({r['age']} {r['sex']}, {r['state']}, "
                f"{int(r['total_runs'])} starts, {int(r['total_wins'])} wins, "
                f"{r['total_earnings']}) — {r['Trainer']}"
            )

    top_lots = df[df["is_top_trainer"]]
    print(f"\n* {len(top_lots)} lots flagged as 'top metro trainer'. "
          f"Curate TOP_METRO_TRAINERS if this looks off.")
    if len(top_lots) > 0:
        print("  Matched trainers in this sale:")
        for trainer in top_lots["Trainer"].drop_duplicates().sort_values():
            count = (top_lots["Trainer"] == trainer).sum()
            print(f"    {count}x  {trainer}")

    raced = df[df["stage"].isin(["MAIDENS", "WINNERS"])]
    neg_days = raced[raced["LS_days_since"].apply(
        lambda x: pd.notna(x) and isinstance(x, (int, float)) and x < 0
    )]
    if len(neg_days) > 0:
        print(f"\n* {len(neg_days)} lots have NEGATIVE Days Since (raced after --date baseline):")
        for _, r in neg_days.iterrows():
            print(
                f"    Lot {int(r['lot']):3d}  {r['horse_name']:20s}  "
                f"LS_date={r['LS_date']}  days_since={r['LS_days_since']}"
            )

    print("=" * 70)
    print()
    return df


# =============================================================================
# CELL VALUE FORMATTING
# =============================================================================
def fmt_cell(val, field):
    if pd.isna(val):
        return "-"
    # v6: Trial_Days_Since added alongside LS_days_since for int coercion
    if field in ("total_runs", "barrier_trial_count", "LS_days_since",
                 "Trial_Days_Since", "rating_display"):
        try:
            return int(val)
        except (ValueError, TypeError):
            return val
    if isinstance(val, str) and val.strip() == "":
        return ""
    return val


# =============================================================================
# VISUAL CUE RESOLVERS — unchanged from v5
# =============================================================================
def get_days_since_fill(days):
    if pd.isna(days) or days in ("-", ""):
        return None
    try:
        d = int(days)
    except (ValueError, TypeError):
        return None
    if d <= 21:
        return DS_GREEN_FILL
    if d <= 60:
        return DS_AMBER_FILL
    return DS_RED_FILL


def get_last_start_fill(ls_value):
    pos = parse_finish_position(ls_value)
    if pos is None:
        return None
    if pos == 1:
        return LS_WIN_FILL
    if pos in (2, 3):
        return LS_PLACE_FILL
    return None


def get_rating_font(rating):
    if pd.isna(rating) or rating in ("-", ""):
        return BODY_FONT
    try:
        r = int(rating)
    except (ValueError, TypeError):
        return BODY_FONT
    if r >= RATING_GREEN_THRESHOLD:
        return RATING_GREEN_FONT
    if r < RATING_GREY_THRESHOLD:
        return RATING_GREY_FONT
    return BODY_FONT


# =============================================================================
# TAB WRITER
# =============================================================================
def write_tab(wb, df, tab_name, stage, col_specs, title_text, subtitle_text):
    sub = df[df["stage"] == stage].sort_values("lot").reset_index(drop=True)
    ws  = wb.create_sheet(tab_name)
    ncols          = len(col_specs)
    last_col_letter = get_column_letter(ncols)

    # Header band
    ws.merge_cells(f"A1:{last_col_letter}1")
    c1 = ws["A1"]
    c1.value     = title_text
    c1.fill      = NAVY_FILL
    c1.font      = TITLE_FONT
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"A2:{last_col_letter}2")
    c2 = ws["A2"]
    c2.value     = subtitle_text
    c2.fill      = NAVY_FILL
    c2.font      = SUBTITLE_FONT
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    ws.row_dimensions[3].height = 6

    for idx, (label, _, _, width) in enumerate(col_specs, start=1):
        cell            = ws.cell(row=4, column=idx, value=label)
        cell.fill       = NAVY_FILL
        cell.font       = HEADER_FONT
        cell.alignment  = CENTER
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[4].height = 28

    field_to_col = {field: i + 1 for i, (_, field, _, _) in enumerate(col_specs)}

    for i, row in sub.iterrows():
        excel_row = 5 + i
        ws.row_dimensions[excel_row].height = 24

        for col_idx, (label, field, alignment, _) in enumerate(col_specs, start=1):
            cell    = ws.cell(row=excel_row, column=col_idx)
            raw_val = row[field] if field in row.index else pd.NA
            display_val = fmt_cell(raw_val, field)

            applied_fill = WHITE_FILL

            # Days Since fill — applies to both raced (LS_days_since) and
            # unraced (Trial_Days_Since) tabs
            if field in ("LS_days_since", "Trial_Days_Since"):
                ds_fill = get_days_since_fill(raw_val)
                if ds_fill:
                    applied_fill = ds_fill

            # Last Start / Last Trial fill — same win/place logic for both
            elif field in ("LS_position", "Last_Trial"):
                ls_fill = get_last_start_fill(raw_val)
                if ls_fill:
                    applied_fill = ls_fill

            cell.fill   = applied_fill
            cell.border = ROW_BORDER

            if field == "lot":
                cell.value = int(raw_val)
                url = row.get("url")
                if pd.notna(url):
                    cell.hyperlink = Hyperlink(ref=cell.coordinate, target=str(url))
                    cell.font = HYPERLINK_FONT
                else:
                    cell.font = BODY_FONT
            elif field == "reserve" and display_val == "Unreserved":
                cell.value = display_val
                cell.font  = BODY_BOLD_FONT
            elif field == "Trainer" and row.get("is_top_trainer"):
                cell.value = f"{TRAINER_STAR}{display_val}"
                cell.font  = TRAINER_STAR_FONT
            elif field == "rating_display":
                cell.value = display_val
                cell.font  = get_rating_font(raw_val)
            else:
                cell.value = display_val
                cell.font  = BODY_FONT

            cell.alignment = alignment

    # Earnings data bar (Winners only)
    if stage == "WINNERS" and "total_earnings" in field_to_col:
        earnings_col_idx    = field_to_col["total_earnings"]
        earnings_col_letter = get_column_letter(earnings_col_idx)
        first_data_row      = 5
        last_data_row       = 5 + len(sub) - 1

        for r_idx in range(len(sub)):
            cell = ws.cell(row=first_data_row + r_idx, column=earnings_col_idx)
            num  = _to_number(cell.value)
            if num is not None:
                cell.value         = num
                cell.number_format = '"A$"#,##0'

        bar_range = f"{earnings_col_letter}{first_data_row}:{earnings_col_letter}{last_data_row}"
        rule = DataBarRule(
            start_type="min", end_type="max",
            color=EARNINGS_BAR_COLOR,
            showValue=True, minLength=0, maxLength=100,
        )
        ws.conditional_formatting.add(bar_range, rule)

    # Footer
    footer_row = 5 + len(sub) + 1
    span_to    = min(5, ncols)
    ws.merge_cells(
        start_row=footer_row, start_column=1,
        end_row=footer_row,   end_column=span_to,
    )
    fc            = ws.cell(row=footer_row, column=1)
    fc.value      = f"Racing Insights  |  racing-insights.com  |  {len(sub)} lots prepared"
    fc.font       = FOOTER_FONT
    fc.alignment  = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[footer_row].height = 20

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{last_col_letter}{4 + len(sub)}"
    ws.sheet_view.showGridLines = False

    print(f"  {tab_name}: {len(sub)} lots, {ncols} cols")


def _to_number(v):
    if v is None or v == "" or v == "-":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[A-Za-z$,\s]", "", str(v).strip())
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


# =============================================================================
# INTRO / COVER PAGE
# =============================================================================
def write_intro(wb, df, sale_title, catalogue_ref, prepared_date):
    """Build the cover tab: logo, sale meta, welcome, contents, legend, about."""
    ws = wb.create_sheet("INTRO", 0)
    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 2

    HEADER_BAND_ROWS = list(range(1, 10))
    for r in HEADER_BAND_ROWS:
        ws.row_dimensions[r].height = 18
        for col_idx in range(1, 8):
            ws.cell(row=r, column=col_idx).fill = NAVY_FILL

    logo_path = Path(LOGO_FILE)
    if logo_path.exists():
        img = XLImage(str(logo_path))
        target_height = 110
        aspect = img.width / img.height
        img.height = target_height
        img.width  = int(target_height * aspect)
        img.anchor = "B2"
        ws.add_image(img)
    else:
        ws.merge_cells("B2:G3")
        c       = ws["B2"]
        c.value = "RACING INSIGHTS"
        c.font  = Font(name=FONT_NAME, size=24, bold=True, color=MINT)
        c.fill  = NAVY_FILL
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for col_idx in range(1, 8):
        ws.cell(row=10, column=col_idx).fill = MINT_FILL
    ws.row_dimensions[10].height = 4

    ws.row_dimensions[11].height = 16
    ws.merge_cells("B12:F12")
    c           = ws["B12"]
    c.value     = sale_title
    c.font      = Font(name=FONT_NAME, size=22, bold=True, color=CHARCOAL_NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[12].height = 32

    ws.merge_cells("B13:F13")
    c       = ws["B13"]
    c.value = (
        f"Catalogue Reference {catalogue_ref}    |    "
        f"Prepared {prepared_date}    |    "
        f"{len(df)} lots"
    )
    c.font      = Font(name=FONT_NAME, size=11, italic=True, color=SOFT_GREY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[13].height = 18

    ws.row_dimensions[14].height = 12

    section_heading_font = Font(name=FONT_NAME, size=13, bold=True, color=CHARCOAL_NAVY)
    body_font = Font(name=FONT_NAME, size=11, color=CHARCOAL_NAVY)

    def add_section_heading(row, title):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        c       = ws.cell(row=row, column=2)
        c.value = title
        c.font  = section_heading_font
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 26
        for col_idx in range(2, 7):
            ws.cell(row=row + 1, column=col_idx).fill = MINT_FILL
        ws.row_dimensions[row + 1].height = 2
        return row + 2

    def add_paragraph(row, text, font=None, height=44):
        font = font or body_font
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        c           = ws.cell(row=row, column=2)
        c.value     = text
        c.font      = font
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = height
        return row + 1

    def add_spacer(row, height=14):
        ws.row_dimensions[row].height = height
        return row + 1

    row = 15
    row = add_section_heading(row, "Welcome")
    row = add_paragraph(
        row,
        "This guide compiles the lots from the current Inglis Digital online sale "
        "that are most relevant to trainers looking to buy a horse to race."
    )
    row = add_paragraph(
        row,
        "Use the tabs at the bottom to browse by stage — Unraced, Maidens, or Winners. "
        "Each tab is sortable and filterable. Click any lot number to open the full "
        "Inglis listing in your browser."
    )
    row = add_spacer(row, 18)

    row = add_section_heading(row, "What you\u2019ll find")
    unraced_n = (df["stage"] == "UNRACED").sum()
    maidens_n = (df["stage"] == "MAIDENS").sum()
    winners_n = (df["stage"] == "WINNERS").sum()
    total_n   = len(df)

    contents_items = [
        ("UNRACED", unraced_n, "Untried prospects \u2014 pedigree and trial form only"),
        ("MAIDENS", maidens_n, "Raced but yet to win \u2014 recent form plus last start"),
        ("WINNERS", winners_n, "Proven on the track \u2014 full record, rating, last start"),
    ]
    for tab_name, count, desc in contents_items:
        ws.cell(row=row, column=2).value     = tab_name
        ws.cell(row=row, column=2).font      = Font(name=FONT_NAME, size=11, bold=True, color=CHARCOAL_NAVY)
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=row, column=3).value     = f"{count} lots"
        ws.cell(row=row, column=3).font      = Font(name=FONT_NAME, size=11, color=MINT, bold=True)
        ws.cell(row=row, column=3).alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
        ws.cell(row=row, column=4).value     = desc
        ws.cell(row=row, column=4).font      = body_font
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1

    ws.cell(row=row, column=2).value = "Total"
    ws.cell(row=row, column=2).font  = Font(name=FONT_NAME, size=11, bold=True, color=CHARCOAL_NAVY)
    ws.cell(row=row, column=3).value = f"{total_n} lots"
    ws.cell(row=row, column=3).font  = Font(name=FONT_NAME, size=11, bold=True, color=CHARCOAL_NAVY)
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
    ws.cell(row=row, column=4).value = "Full racing horses, all eligible categories"
    ws.cell(row=row, column=4).font  = Font(name=FONT_NAME, size=11, italic=True, color=SOFT_GREY)
    ws.row_dimensions[row].height = 22
    row += 1
    row = add_spacer(row, 18)

    row = add_section_heading(row, "How to read the colour cues")

    def add_legend_row(row, sample_text, sample_font, sample_fill, description):
        sample_cell           = ws.cell(row=row, column=2)
        sample_cell.value     = sample_text
        sample_cell.font      = sample_font
        if sample_fill:
            sample_cell.fill  = sample_fill
        sample_cell.alignment = Alignment(horizontal="center", vertical="center")
        sample_cell.border    = ROW_BORDER
        desc_cell             = ws.cell(row=row, column=3)
        desc_cell.value       = description
        desc_cell.font        = body_font
        desc_cell.alignment   = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        ws.row_dimensions[row].height = 22
        return row + 1

    def add_legend_subheading(row, text):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        c           = ws.cell(row=row, column=2)
        c.value     = text
        c.font      = Font(name=FONT_NAME, size=11, bold=True, color=CHARCOAL_NAVY, italic=True)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 20
        return row + 1

    row = add_legend_subheading(row, "Days since last run")
    row = add_legend_row(row, "12", body_font, DS_GREEN_FILL, "In work \u2014 21 days or less")
    row = add_legend_row(row, "45", body_font, DS_AMBER_FILL, "Let-up \u2014 22 to 60 days")
    row = add_legend_row(row, "90", body_font, DS_RED_FILL,   "Long spell \u2014 over 60 days, needs preparation")
    row = add_spacer(row, 8)

    row = add_legend_subheading(row, "Last start / last trial finish")
    row = add_legend_row(row, "1st of 9", body_font, LS_WIN_FILL,   "Won")
    row = add_legend_row(row, "3rd of 9", body_font, LS_PLACE_FILL, "Placed (2nd or 3rd)")
    row = add_legend_row(row, "7th of 9", body_font, None,          "Unplaced \u2014 no fill applied")
    row = add_spacer(row, 8)

    row = add_legend_subheading(row, "Rating (Winners tab)")
    row = add_legend_row(row, "78", RATING_GREEN_FONT, None,
                         f"Top tier for this sale ({RATING_GREEN_THRESHOLD}+)")
    row = add_legend_row(row, "65", body_font, None,
                         f"Mid tier ({RATING_GREY_THRESHOLD}\u2013{RATING_GREEN_THRESHOLD - 1})")
    row = add_legend_row(row, "55", RATING_GREY_FONT, None,
                         f"Lower country grade (under {RATING_GREY_THRESHOLD})")
    row = add_spacer(row, 8)

    row = add_legend_subheading(row, "Trainer, earnings, and other markers")
    row = add_legend_row(
        row, "\u2605 Trainer", TRAINER_STAR_FONT, None,
        "Mint star \u2014 recognised top metropolitan trainer"
    )
    row = add_legend_row(
        row, "\u25ae\u25ae\u25ae\u25ae",
        Font(name=FONT_NAME, size=10, bold=True, color=MINT), None,
        "Earnings bar (Winners tab) \u2014 length scales with total career prizemoney"
    )
    row = add_legend_row(row, "Unreserved", BODY_BOLD_FONT, None,
                         "Bold text \u2014 no reserve price set, sale is more likely")
    row = add_legend_row(row, "Lot #", HYPERLINK_FONT, None,
                         "Click the lot number to open the full Inglis Digital listing")
    row = add_legend_row(row, REPORT_PRESENT, body_font, None,
                         "Report attached on the Inglis listing \u2014 worth reviewing")
    row = add_spacer(row, 18)

    row = add_section_heading(row, "About Racing Insights")
    row = add_paragraph(
        row,
        "Racing Insights compiles enriched sale catalogues for thoroughbred buyers \u2014 "
        "compressing the discovery phase so serious buyers can spend their time appraising "
        "horses, not trawling for them.",
    )
    row = add_paragraph(
        row,
        "Your Stable, Backed by Data    \u00b7    racing-insights.com",
        font=Font(name=FONT_NAME, size=11, italic=True, color=MINT, bold=True),
        height=22,
    )


# =============================================================================
# MAIN
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description=(
            "Build the Racing Insights Buyers\u2019 Guide from the merged Excel "
            "produced by merge_and_split.py."
        )
    )
    parser.add_argument("--sale", required=True,
                        help="Sale code, e.g. 2026-05B")
    parser.add_argument("--date", required=True,
                        help="Prepared date YYYY-MM-DD, e.g. 2026-06-06")
    args = parser.parse_args()

    try:
        datetime.strptime(args.date, "%Y-%m-%d")
    except ValueError:
        sys.exit(f"❌ --date must be YYYY-MM-DD, got: {args.date!r}")

    sale         = args.sale
    date_str     = args.date
    sale_title   = sale_title_from_code(sale)
    catalogue_ref = sale
    prepared_date = prepared_date_from_str(date_str)

    input_path  = Path(f"{sale}_merged.xlsx")
    output_path = Path(f"BuyersGuide_{sale}_{date_str}.xlsx")

    if not input_path.exists():
        sys.exit(
            f"❌ Input not found: {input_path}\n"
            f"   Run merge_and_split.py --sale {sale} --date {date_str} first."
        )

    print(f"Reading: {input_path}")
    df = preprocess(input_path)

    subtitle = (
        f"Catalogue Reference: {catalogue_ref}  |  "
        f"Prepared: {prepared_date}  |  "
        f"Racing Insights \u2014 Your Stable, Backed by Data"
    )

    wb = Workbook()
    wb.remove(wb.active)

    print("Building tabs:")
    write_tab(wb, df, "UNRACED", "UNRACED", UNRACED_COLS, sale_title, subtitle)
    write_tab(wb, df, "MAIDENS", "MAIDENS", MAIDENS_COLS, sale_title, subtitle)
    write_tab(wb, df, "WINNERS", "WINNERS", WINNERS_COLS, sale_title, subtitle)
    write_intro(wb, df, sale_title=sale_title,
                catalogue_ref=catalogue_ref, prepared_date=prepared_date)

    wb.save(output_path)
    print(f"\nSaved: {output_path.resolve()}")


if __name__ == "__main__":
    main()
