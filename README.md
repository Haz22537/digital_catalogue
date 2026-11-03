# Racing Insights — Buyers' Guide Pipeline

Per-sale data pipeline for the Inglis Digital Buyers' Guide.

**Last major change: 22 August 2026 — API pivot.** Inglis Digital's site is a
thin SPA over a public JSON API. Playwright scraping of Inglis is retired. See
`RacingInsights_APIPivot_Handover_22Aug2026.md` for the full background,
endpoint reference, and the list of gotchas that will otherwise bite twice.

Most scripts take `--sale` in short form (`2026-08B`, A=early, B=late). Scripts
that touch the archive also accept `--slug` (`2026-august-late-online-sale`),
which is required for single-sale months — see *Sale codes* below.

---

## Setup

### Prerequisites

Python 3.8+, with:

```
pip install requests beautifulsoup4 pandas openpyxl selenium
```

**Selenium** is needed only for `02_RA_url_discovery.py` — RA's website has
dynamic content and requires a browser. Everything else is plain HTTP.

For Selenium, download a **ChromeDriver** matching your Chrome version from
https://chromedriver.chromium.org and place it in your PATH or the project
root. The script looks for `chromedriver` (macOS/Linux) or `chromedriver.exe`
(Windows).

### Project structure

```
inglis_digital/
├── *.py                           All pipeline scripts
├── sale_codes.py                  Path + short-code helpers
├── data/
│   └── raw/
│       ├── sales/{slug}/          Downloaded sale archives
│       └── pedigree/              Global pedigree store
├── {sale}_stats.csv               Outputs from each step
├── {sale}_RA_*.csv
├── {sale}_merged.xlsx
└── {sale}_failures.csv
```

Create the `data/` directory before the first run, or the scripts will create
it automatically.

### Your first sale

```bash
python inglis_api_scrape.py --sale 2026-08B --slug 2026-august-late-online-sale
python build_stats_from_api.py --sale 2026-08B --slug 2026-august-late-online-sale
```

Both commands need the slug because the API keys on it, not the short code.
After that, all subsequent scripts use the short code alone since they work
from disk.

Check for the freshness warning after `build_stats_from_api.py`. If anything is
flagged, see "Two checks that are not optional" below.

---

## Per-sale workflow

| # | Step | Command | Output |
|---|------|---------|--------|
| 1 | Fetch sale from API | `python inglis_api_scrape.py --sale 2026-august-late-online-sale` | `data/raw/sales/{slug}/`, `data/raw/pedigree/` |
| 2 | Archive → stats | `python build_stats_from_api.py --sale 2026-08B --slug {slug}` | `2026-08B_stats.csv` |
| 3 | Split manifests | `python manifest_convert.py --sale 2026-08B` | `_manifest_RA.csv`, `_manifest_RA_unraced.csv` |
| 4 | RA URL + dam check | `python 02_RA_url_discovery.py --sale 2026-08B` | `_RA_urls.csv` |
| 5 | **Review dam mismatches** (see below) | — | edits saved in place |
| 6 | RA unraced trials | `python 02b_RA_unraced_trials.py --sale 2026-08B` | `_RA_unraced_trials.csv` |
| 7 | RA ratings + trainer | `python 03_RA_ratings_trainer.py --sale 2026-08B` | `_RA_ratings.csv` |
| 8 | Merge + split | `python merge_and_split.py --sale 2026-08B --date 2026-08-22` | `_merged.xlsx` |
| 9 | Build guide | `python build_buyers_guide_v6.py --sale 2026-08B --date 2026-08-22` | Buyers' Guide xlsx |

Errors at any step append to `2026-08B_failures.csv`.

**Steps 1–2 replace ~40 minutes of Playwright plus a manual HTML save with
about 5 minutes of HTTP.** RA (steps 4–7) is now the only slow part.

### Alternative to step 3

`make_ra_manifests.py` builds the same two manifests directly from the archive,
skipping `stats.csv`. Useful when you want RA started before the stats parse:

```
python make_ra_manifests.py --sale 2026-08B --slug 2026-august-late-online-sale
```

It also writes `{sale}_ra_split_evidence.csv` — one row per lot showing the
raced/unraced decision, which signal produced it, and record freshness.

---

## Two checks that are not optional

### Check the freshness warning after step 2

Arion re-extracts a catalogue **progressively**, and there is no signal in the
API saying "extraction complete". Scrape while that job is running and some
records are pre-refresh.

This happened on 2026-08B: 24 of 160 lots were stale, and six horses with real
race records read `Unraced.` — they went onto the unraced manifest and would
have appeared on the wrong tab of a published guide.

Both `build_stats_from_api.py` and `make_ra_manifests.py` compare every
record's `raceRecordUpdatedAt` against the sale's `catalogueReleaseDate` and
warn. If anything is flagged:

```
python inglis_api_scrape.py --sale {slug} --refresh-stale
```

Then re-run from step 2. `--refresh-stale` refetches only the stale records,
which is far cheaper than `--force`.

**Best avoided entirely: scrape a day or two after catalogue release, not the
morning bidding opens.**

### Step 5 — dam mismatch review

Step 4 writes a `dam_match` column on every row:

| `dam_match` | Meaning | Action |
|---|---|---|
| `auto` | RA dam matches manifest dam. Right horse, almost certainly. | None — step 7 will process. |
| `mismatch` | RA returned a different dam. Probably the wrong horse. | Investigate and fix. |
| `none` | No dam extracted (no result, or error). | Same as mismatch. |

In Excel: open `{sale}_RA_urls.csv`, filter `dam_match != "auto"`, and for each
row either replace `ra_profile_url` with the correct one or leave it (step 7
skips and logs it). Set `dam_match` to `auto` on rows you've fixed, save, then
run step 7.

Typically fewer than 5 rows per sale, mostly dam names that changed on import —
2–3 minutes of work. Step 7 only processes `auto` rows, so anything you don't
get to stays out of the final data rather than silently corrupting it.

---

## Why each script exists

| Script | What it earns |
|---|---|
| `inglis_api_scrape.py` | Fetches a whole sale from the public API into a raw gzipped JSON archive. Resumable, append-only fetch log. One request gets the entire catalogue. |
| `build_stats_from_api.py` | Parses the archive into `{sale}_stats.csv` in the legacy schema, so everything downstream runs unmodified. Includes the freshness gate and `--as-of` reconstruction. |
| `make_ra_manifests.py` | Archive → RA manifests directly, bypassing stats. Lets RA start early. Writes the split-evidence file. |
| `manifest_convert.py` | Splits stats into raced/unraced RA manifests, filtered to RH/RF/2YO. Warns on unrecognised category codes. |
| `02_RA_url_discovery.py` | Selenium search of Racing Australia per horse, plus inline dam-match. The most fragile script in the pipeline. |
| `02b_RA_unraced_trials.py` | Trainer, stable and trial detail for unraced horses. **Still required** — API trial rows carry no trainer. |
| `03_RA_ratings_trainer.py` | Plain HTTP fetch of ratings + trainer. NSW dual-rating split happens here. |
| `merge_and_split.py` | Joins stats + RA data, computes Days Since, writes the four-sheet intermediate workbook. |
| `build_buyers_guide_v6.py` | Final formatting and branding. Untouched by the API pivot. |
| `audit_record_freshness.py` | Per-lot stored-vs-live comparison. Run when you want certainty rather than a local heuristic. |
| `probe_racerecord.py` | Parser regression harness over the whole local archive. Re-run after any parser change. |

### Retired — do not resurrect

`01_catalogue_scrape.py`, `01b_catalogue_from_html.py`,
`inglisdigital_lotscraper.py`.

All three scraped the Inglis DOM. Between them they carried two live bugs that
the API design makes structurally impossible: share lots were stubbed out with
zeros, and unnamed lots collided in a name-keyed cache. The API keys on
`pedigreeId`, which is unique per horse and present even when the horse has no
name.

---

## Data quirks the pipeline handles

- **NSW dual benchmarks.** Racing NSW reports two ratings comma-separated,
  e.g. `75,62` = Metro/Provincial 75, Country 62. Step 7 splits these into
  `Rating` and `Rating_NSW_Country`. Other states have one value and
  `Rating_NSW_Country` stays blank. A literal `0` means raced 0 times in that
  sub-jurisdiction.
- **Blank ratings on maidens are correct.** Maiden races place horses on win
  record, not rating. Recent winners awaiting a BM assignment, and
  first-start-in-jurisdiction horses, legitimately show blank. Backfill
  manually from the previous rating.
- **Blank vs zero on `total_runs`.** Unraced horses must get **blank**, not `0`.
  `merge_and_split.load_stats()` coerces to numeric so blanks become `NaN`, and
  `classify_stage()` keys the UNRACED tab off `NaN`. Writing `0` puts every
  unraced horse on the MAIDENS tab.
- **Arion markup quirks.** Duplicate `class` attributes
  (`<table class="rrTable" class="talt">`) make `html.parser` drop the first
  value, hiding the table entirely. Every parser needs the `RE_DUP_CLASS` fix.
  Also: a literal `n` appears between some tags, and finish positions include
  non-numeric codes (`LR`) and dead heats (`5th=`).
- **Share lots.** A horse can appear as multiple RS lots at different
  percentages. `shareToSell` in the catalogue gives the percentage directly, and
  `pedigreeId` dedupes the underlying horse. RS is excluded from the guide by
  `RACEABLE_CATEGORIES` but the data is captured.
- **Same-name horses across jurisdictions.** Handled by the dam-match gate.
- **Broodmare foal-at-foot seasonality is real, not a data fault.** Australian
  mares foal Aug–Dec; by late July foals are weaned. A 0% July reading against
  16% in August is expected.

---

## Resume after a crash

All long-running scripts write incrementally and flush after every row. Re-run
the same command — anything already in the output is skipped.

- `inglis_api_scrape.py` skips pedigree records already on disk. Use `--force`
  to refetch everything, or `--refresh-stale` for just the stale ones.
- `02`, `02b` and `03` skip lots already present in their own output.
- To retry a single row, delete it from the output CSV first.

**If a manifest changes after RA has run** (for example after a
`--refresh-stale` moves horses from unraced to raced), re-running `02` and `03`
picks up only the new lots — but `02b`'s output still contains the horses that
are no longer unraced. Remove those rows manually, or they will appear on two
tabs.

---

## Sale codes

Canonical short form: `2026-08B`. The Inglis slug
`2026-august-late-online-sale` is derived only when constructing an API call.
See `sale_codes.py`.

**Known limitation:** `short_to_slug()` assumes an A/B suffix and cannot express
single-sale months. January 2026 was one sale
(`2026-january-online-sale`) with no A/B. Pass `--slug` explicitly for those.

---

## Files per sale (`--sale 2026-08B`)

```
data/raw/sales/2026-august-late-online-sale/
    auction_info.json.gz          sale metadata
    catalogue.json.gz             closed sales — immutable, fixed name
    catalogue__20260821T090412.json.gz   open sales — timestamped snapshots
    fetch_log.csv                 append-only request log

data/raw/pedigree/
    {pedigreeId}.json.gz          global store, shared across all sales

2026-08B_stats.csv                step 2 output
2026-08B_manifest_RA.csv          step 3 output / step 4 input
2026-08B_manifest_RA_unraced.csv  step 3 output / step 6 input
2026-08B_ra_split_evidence.csv    raced/unraced decisions + freshness
2026-08B_RA_urls.csv              step 4 output — edit in place
2026-08B_RA_unraced_trials.csv    step 6 output
2026-08B_RA_ratings.csv           step 7 output
2026-08B_merged.xlsx              step 8 output
2026-08B_failures.csv             aggregated errors
```

The pedigree store is **global, not per-sale**. A horse relisted across sales is
fetched once. This is why `--refresh-stale` exists: without it, a record fetched
long ago is never revisited.

---

## Debugging

**API unreachable?** `python api_check.py` tests the endpoints directly and
reports header sensitivity, so you can tell a network problem from a change at
Inglis's end.

**Parser behaving oddly?** `python probe_racerecord.py` runs every candidate
pattern across the whole local archive (~7,700 records) and reports match rates
and reconciliation. Far better evidence than one sale can give.

**A horse on the wrong tab?** Check `{sale}_ra_split_evidence.csv` for its
decision and evidence string, then `audit_record_freshness.py` for a live
comparison.

---

## Known gaps

| Field | Status |
|---|---|
| `disclosures` | Only on the lot-entry endpoint, which is never called (see handover §2). `animalDefects` requested from Inglis directly. |
| `gst` | Same. |
| `report_count`, `reports` | The `all-attachments` endpoint is unprobed. Hardcoded `0`. One call would settle it. |

---

## Roadmap

- **DuckDB loader.** ~9,600 lots across 26 closed sales are captured but not
  loaded. Scrapers write raw; only the loader writes to the database.
- **Price model.** Filter `lotStatus == 'Sold'` before touching `currentBid` —
  it reads `0`, not null, when no bid was placed. Split train/test **across sale
  cohorts**, not randomly: clearance rate swings ~10 points between sales.
- **Days Since display bugs.** Zeros render as dashes and aren't colour-coded.
- **`TOP_METRO_TRAINERS`** needs aligning with real-world top stables.
- **`Rating_Juristiction`** typo, preserved from Power Query, to be corrected
  across scripts once the pipeline is stable.
