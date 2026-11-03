# -*- coding: utf-8 -*-
"""
Pipeline-wide utilities — failure logging and shared helpers.

The failure log is a single CSV per sale that aggregates errors from
every phase (catalogue scrape, lot scrape, RA URL discovery, RA ratings).
One file makes triage simple: open it, sort by phase or status, you see
everything that went wrong in one place.
"""

import csv
import re
from datetime import datetime
from pathlib import Path

from sale_codes import failures_path


FAILURE_FIELDS = [
    "timestamp", "phase", "lot", "horse_name", "status", "fail_reason",
]

# Inglis Digital now explicitly suffixes Australian-bred horses with "(AUS)";
# previously the absence of a suffix implied AUS. Racing Australia never shows
# a country code for AUS-bred horses — only for imports. So "(AUS)" must be
# stripped before comparing, while all other suffixes (NZ, GB, USA, IRE, ...)
# are left intact because both sources agree on those.
_AUS_SUFFIX_RE = re.compile(r"\s*\(\s*AUS\s*\)\s*$", re.IGNORECASE)


def log_failure(
    sale: str,
    phase: str,
    *,
    lot: str | int = "",
    horse_name: str = "",
    status: str = "",
    fail_reason: str = "",
) -> None:
    """Append one failure row to {sale}_failures.csv. Creates file + header on first call."""
    path = failures_path(sale)
    write_header = not Path(path).exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FAILURE_FIELDS, extrasaction="ignore")
        if write_header:
            writer.writeheader()
        writer.writerow({
            "timestamp":   datetime.now().isoformat(timespec="seconds"),
            "phase":       phase,
            "lot":         str(lot) if lot != "" else "",
            "horse_name":  horse_name or "",
            "status":      status or "",
            "fail_reason": (fail_reason or "")[:500],  # cap free-text length
        })


def normalize_dam(name: str | None) -> str:
    """Case-fold + whitespace-collapse for dam-name comparison.

    A trailing '(AUS)' suffix IS stripped: Inglis Digital started tagging
    Australian-bred horses explicitly, but Racing Australia only ever shows
    country codes for imports, so the suffix is a formatting difference rather
    than a genuine mismatch.

    All other country suffixes (e.g. '(NZ)', '(USA)') are intentionally NOT
    stripped — Inglis and RA agree on those, so stripping would mask genuine
    mismatches.
    """
    cleaned = _AUS_SUFFIX_RE.sub("", (name or "").strip())
    return re.sub(r"\s+", " ", cleaned.strip().lower())
