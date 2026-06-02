# Racing Insights — RA Pipeline

Per-sale data-gathering pipeline for the Inglis Digital Buyer's Guide.
All scripts take a single `--sale` argument in `YYYY-MMD` form (A=early, B=late).

## Per-sale workflow

| # | Step                       | Command                                              | Output                          |
|---|----------------------------|------------------------------------------------------|---------------------------------|
| 1 | Build manifest from catalogue | `python 01_catalogue_scrape.py --sale 2026-05B`   | `2026-05B_manifest.csv`         |
| 2 | Scrape Inglis lot pages    | `python inglisdigital_lotscraper.py --sale 2026-05B` | `_stats.csv`, `_text.csv`       |
| 3 | Filter manifest for RA     | `python manifest_convert.py --sale 2026-05B`         | `_manifest_RA.csv`              |
| 4 | Phase 1: RA URL + dam check| `python 02_RA_url_discovery.py --sale 2026-05B`      | `_RA_urls.csv`                  |
| 5 | **Review mismatches in Excel** (see below)                                       | edits saved in place            |
| 6 | Phase 2: ratings + trainer | `python 03_RA_ratings_trainer.py --sale 2026-05B`    | `_RA_ratings.csv`               |

Any errors at any step append to `2026-05B_failures.csv`.

## Step 5 — what you do

Phase 1 writes a `dam_match` column on every row:

| `dam_match` | Meaning                                              | Action                       |
|-------------|------------------------------------------------------|------------------------------|
| `auto`      | RA dam matches manifest dam. Right horse, almost certainly. | None — Phase 2 will process. |
| `mismatch`  | RA returned a different dam. Probably wrong horse.   | Investigate & fix, see below |
| `none`      | Scrape returned no dam (no result, or error).        | Same as mismatch.            |

In Excel:
1. Open `2026-05B_RA_urls.csv`.
2. Filter `dam_match != "auto"` — typically <10% of rows.
3. For each: either replace `ra_profile_url` with the correct one (find it manually on RA), or leave it (Phase 2 will skip and log).
4. Change `dam_match` to `auto` on the rows you've fixed.
5. Save. Run Phase 2.

Phase 2 will only process rows where `dam_match == auto`, so anything you don't get to stays out of the final data — no silent corruption.

## Why each script exists

| Script                          | What it does that earns its existence                                                                 |
|---------------------------------|-------------------------------------------------------------------------------------------------------|
| `01_catalogue_scrape.py`        | Replaces hand-transcribing 150+ rows from the Inglis catalogue page. Pure time saver.                 |
| `inglisdigital_lotscraper.py`   | Captures pedigree, race record, sale-page disclosures from every lot's detail page.                   |
| `manifest_convert.py`           | Drops unraced horses and non-raceable categories before Phase 1. Phase 1 is the slowest step; this cuts wall-clock by 30–50%. |
| `02_RA_url_discovery.py`        | Selenium search of Racing Australia per horse + inline dam-match. The most fragile script — review the gate. |
| `03_RA_ratings_trainer.py`      | Plain HTTP fetch of each profile's ratings page + trainer info. NSW dual-rating split happens here.   |

## Data quirks the pipeline handles

- **NSW dual benchmarks.** Racing NSW reports two ratings comma-separated, e.g. `75,62` = Metro/Provincial 75, Country 62. Phase 2 splits these into `Rating` and `Rating_NSW_Country`. Other states have one value (no comma) and `Rating_NSW_Country` stays blank. A literal `0` means raced 0 times in that sub-jurisdiction.
- **Inglis dash zeros.** Race-record count columns (`total_runs`/`wins`/`2nd`/`3rd`, `age{N}_*`) render `-` on Inglis pages for zero. The lot scraper now writes numeric `0` instead. Earnings columns are untouched (they render `$0.00` rather than `-`).
- **Duplicate share lots.** A horse can appear as multiple RS share lots. The Inglis scraper caches by horse base name so each unique horse is fetched once; share-percentage rows reuse the data.
- **Same-name horses across jurisdictions.** Handled by the dam-match gate above.

## Resume after a crash

Phase 1, Phase 2, and the Inglis lot scraper all write incrementally. Re-run the same command after a crash — anything already in the output is skipped. To force a retry on a single row, delete it from the output CSV first.

The catalogue scraper does not resume — it's a single page load, just re-run it.

## Sale codes

Canonical short form: `2026-05B`. The Inglis URL slug `2026-may-late-online-sale` is derived only when an Inglis URL needs constructing. See `sale_codes.py` for the mapping and filename conventions — change anything there and it propagates.

## Files per sale (using `--sale 2026-05B`)

```
2026-05B_manifest.csv          ← step 1 output (catalogue scrape)
2026-05B_stats.csv             ← step 2 output (structured fields)
2026-05B_text.csv              ← step 2 output (free text)
2026-05B_manifest_RA.csv       ← step 3 output / step 4 input
2026-05B_RA_urls.csv           ← step 4 output (single file, edit in place)
2026-05B_RA_ratings.csv        ← step 6 output (final RA enrichment)
2026-05B_failures.csv          ← aggregated errors from any step
```

## Catalogue scraper — first-run debugging

The catalogue page is a Vue.js SPA, so its DOM changes whenever Inglis re-skin. If `01_catalogue_scrape.py` returns 0 lots or obviously wrong data, run:

```
python 01_catalogue_scrape.py --sale 2026-05B --snapshot
```

This writes `2026-05B_catalogue_snapshot.html` and `.png`. Share those for selector tuning.

## Future (Tier 3+)

- **Cross-sale database.** A loader that ingests `*_stats.csv`, `*_text.csv`, and `*_RA_ratings.csv` into DuckDB/SQLite. Schemas are already stable across sales to support this. Enables: sire/dam-sire/trainer/vendor profiling, trend analysis across sales, win-rate aggregation. Build when the per-sale workflow feels routine.
- **Post-sale results.** Sale price + buyer info aren't on any current page scraped — they live on the post-sale results page. A separate script reading that page and joining on `lot_id` would complete the cross-sale picture.
- **Buyer's Guide formatter.** Replace the manual Excel join with a script that segments by performance band (Unraced / Maidens / Winners / Earners) and writes the formatted workbook.
