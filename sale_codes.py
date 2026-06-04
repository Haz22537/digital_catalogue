"""
Sale-code conventions and filename derivation.

Two forms of sale identifier are in use across the pipeline:

    short form    "2026-05B"                       canonical, human-friendly
    URL slug      "2026-may-late-online-sale"      what Inglis URLs require

All scripts accept the short form via --sale. The URL slug is derived only
when a script needs to construct an Inglis URL. Keeping the mapping in one
place means a typo or convention tweak is a one-file change.
"""

MONTHS = {
    "01": "january", "02": "february", "03": "march",     "04": "april",
    "05": "may",     "06": "june",     "07": "july",      "08": "august",
    "09": "september","10": "october", "11": "november",  "12": "december",
}
MONTH_TO_NUM = {name: num for num, name in MONTHS.items()}
TIMING_WORDS = {"A": "early", "B": "late"}
TIMING_LETTERS = {word: letter for letter, word in TIMING_WORDS.items()}


def short_to_slug(short: str) -> str:
    """'2026-05B' -> '2026-may-late-online-sale'"""
    try:
        year, monthcode = short.split("-")
        month_num, timing = monthcode[:2], monthcode[2:].upper()
    except (ValueError, IndexError):
        raise ValueError(
            f"Invalid sale code {short!r} — expected form like '2026-05B'"
        )

    if month_num not in MONTHS:
        raise ValueError(f"Unknown month '{month_num}' in {short!r}")
    if timing not in TIMING_WORDS:
        raise ValueError(
            f"Unknown timing '{timing}' in {short!r} (expected A=early or B=late)"
        )

    return f"{year}-{MONTHS[month_num]}-{TIMING_WORDS[timing]}-online-sale"


def slug_to_short(slug: str) -> str:
    """'2026-may-late-online-sale' -> '2026-05B'. Falls back to slug if unparsable."""
    parts = slug.lower().split("-")
    year   = next((p for p in parts if p.isdigit() and len(p) == 4), None)
    month  = next((MONTH_TO_NUM[p] for p in parts if p in MONTH_TO_NUM), None)
    timing = next((TIMING_LETTERS[p] for p in parts if p in TIMING_LETTERS), "")
    if year and month:
        return f"{year}-{month}{timing}"
    return slug


# ─── Filename conventions ──────────────────────────────────────────────
# Single source of truth so all six scripts agree on what files exist
# where. Change a convention here and it propagates everywhere.

def manifest_path(sale: str) -> str:
    """Initial sale manifest (input to Inglis scraper). Hand-built today."""
    return f"{sale}_manifest.csv"

def stats_path(sale: str) -> str:
    """Inglis scraper output — structured fields."""
    return f"{sale}_stats.csv"

def text_path(sale: str) -> str:
    """Inglis scraper output — free-text fields."""
    return f"{sale}_text.csv"

def manifest_RA_path(sale: str) -> str:
    """Filtered manifest from manifest_convert.py — Phase 1 input."""
    return f"{sale}_manifest_RA.csv"

def manifest_RA_unraced_path(sale: str) -> str:
    """Unraced filtered manifest from manifest_convert.py — Phase 1b input."""
    return f"{sale}_manifest_RA_unraced.csv"

def ra_unraced_trials_path(sale: str) -> str:
    """Phase 1b output — trainer + trial data for unraced horses."""
    return f"{sale}_RA_unraced_trials.csv"

def ra_urls_path(sale: str) -> str:
    """Phase 1 output — horse URLs from Racing Australia."""
    return f"{sale}_RA_urls.csv"

def ra_urls_path(sale: str) -> str:
    """Phase 1 output — horse URLs from Racing Australia."""
    return f"{sale}_RA_urls.csv"

def ra_ratings_path(sale: str) -> str:
    """Phase 2 output — final RA-enriched data."""
    return f"{sale}_RA_ratings.csv"

def failures_path(sale: str) -> str:
    """Aggregated failure log across all pipeline phases."""
    return f"{sale}_failures.csv"

def catalogue_snapshot_path(sale: str) -> str:
    """HTML snapshot of the Inglis catalogue page for selector debugging."""
    return f"{sale}_catalogue_snapshot.html"
