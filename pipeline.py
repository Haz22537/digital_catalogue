# -*- coding: utf-8 -*-
"""
Pipeline-wide utilities — failure logging and shared helpers.

The failure log is a single CSV per sale that aggregates errors from
every phase (catalogue scrape, lot scrape, RA URL discovery, RA ratings).
One file makes triage simple: open it, sort by phase or status, you see
everything that went wrong in one place.
"""

import csv
from datetime import datetime
from pathlib import Path

from sale_codes import failures_path


FAILURE_FIELDS = [
    "timestamp", "phase", "lot", "horse_name", "status", "fail_reason",
]


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

    Country suffixes (e.g. '(USA)') are intentionally NOT stripped — Inglis and
    RA agree on suffix presence, so stripping would mask genuine mismatches.
    """
    import re
    return re.sub(r"\s+", " ", (name or "").strip().lower())
