# SEC EDGAR Research & Financial Modeling Toolkit

A Jupyter-notebook-based tool that pulls real financial data and 10-K filing text directly from the SEC's public EDGAR system for a given stock ticker, and turns it into a formatted, formula-driven Excel workbook: a three-statement financial model (income statement, balance sheet, cash flow, key ratios) plus optional competitive-intelligence tabs (peer market share, product-type/end-market classification, year-over-year risk-factor changes, and a scored research summary).

Unlike a typical scraper, this pulls structured data from SEC's own APIs (the XBRL "company facts" API and the EDGAR full-text filing archive) rather than parsing arbitrary vendor HTML, and it's built with real attention to SEC's usage rules (a descriptive User-Agent, rate limiting, retry/backoff on 403/429).

## Contents

| File | Purpose |
|---|---|
| `RUN_CODE.ipynb` | The entry point. A single notebook cell that enables `autoreload` and imports `competitor_intel_plugin` — importing that module is what wires up and displays the tool's UI panels in Jupyter. |
| `sec_model_builder.py` | Core SEC data layer: rate-limited/retrying HTTP session, ticker→CIK resolution, XBRL company-facts fetching, the GAAP-tag "concept taxonomy," annual-series extraction, and the Excel three-statement model builder. Also defines its own standalone ipywidgets UI. |
| `filing_extractor.py` | Fetches a company's filing history and profile (SIC code), locates the Nth-most-recent 10-K, downloads and flattens its HTML to plain text, isolates the "Item 1 Business" and "Item 1A Risk Factors" sections, mines candidate competitor names out of the Business section, and looks up SIC-code peers via EDGAR's company-search endpoint. |
| `analysis_engine.py` | Heuristic analysis on the extracted filing text: consumable-vs-durable product classification, end-market keyword detection, year-over-year Risk Factors diffing (via `difflib`), and peer-group market-share computation. |
| `excel_plugins.py` | Builds three additional worksheet tabs on top of the base model: "Competitive Landscape" (peer revenue/share table), "Industry Factors" (product type, end markets, new risk language), and "Research Signal Summary" (a small scored scorecard). |
| `competitor_intel_plugin.py` | The orchestrator: `run_full_research_workbook()` ties every module above together into one pipeline, and defines the notebook's main "Build Full Research Workbook" button UI. |

There is also a `__MACOSX/` folder and `.DS_Store`/`._*` files in the zip — these are macOS Finder metadata created automatically when zipping on a Mac, not part of the project; safe to delete.

## What it actually does, end to end

Running `RUN_CODE.ipynb` and clicking "Build Full Research Workbook" for a ticker (e.g. `ILMN`) does the following, all against live SEC EDGAR data:

1. Resolves the ticker to a CIK and official company name using SEC's public ticker registry (cached locally for a day).
2. Fetches the company's XBRL "company facts" and extracts up to N years of annual figures for ~25 standard line items (revenue, COGS, SG&A, R&D, operating income, tax, net income, EPS, balance sheet items, cash flow items) using a fallback list of GAAP tags per concept, since different filers tag the same concept differently.
3. Builds a formatted Excel workbook with a proper three-statement model — the calculated rows (gross profit, margins, YoY growth, balance check, ratios) are written as **live Excel formulas**, not pre-computed static numbers, so the workbook stays usable/auditable after the fact rather than being a frozen export.
4. Separately, fetches the company's latest and prior 10-K filing documents, strips HTML down to plain text, and isolates the "Item 1 Business" and "Item 1A Risk Factors" sections using a heading-position heuristic designed to skip past the table of contents.
5. Mines the Business section for candidate competitor names (via corporate-suffix pattern matching and via capitalized phrases near words like "compete"/"rival"), then confirms which candidates are real public companies by cross-referencing SEC's own company-title registry — candidates that don't match are kept separately as likely-private competitors rather than discarded.
6. Cross-checks against other companies sharing the same 4-digit SIC code (via EDGAR's `browse-edgar` endpoint) as a second, structured source of peers, in addition to the text-mined ones.
7. Computes each peer's revenue and share of the combined peer group's revenue across the timeline, and diffs this year's Risk Factors section against last year's (via fuzzy sentence matching) to flag likely new risk language.
8. Classifies the business as more "consumable/recurring" or "durable/one-time" leaning, and detects which of ~25 common end-markets are mentioned, using keyword heuristics.
9. Assembles everything into one `.xlsx`: the three-statement model plus "Competitive Landscape," "Industry Factors," and "Research Signal Summary" tabs, then opens it in the OS's default spreadsheet app.

`sec_model_builder.py` also works as a lighter-weight tool on its own — its "Build Model Matrix" button (steps 1–3 above only) produces just the three-statement model without the competitive-intelligence tabs.

## Module details

### `sec_model_builder.py`

- **SEC compliance / request layer** — a shared `requests.Session` with a descriptive `User-Agent` (SEC requires a real contact identifier on automated requests), a small delay between requests to stay under SEC's ~10 requests/second guidance, and `sec_get()`, which retries with exponential backoff specifically on `403`/`429` responses rather than treating a rate-limit block page as "no data."
- **Ticker registry** — `load_ticker_map()` pulls SEC's `company_tickers.json` and caches it to `sec_tickers_cache.json` in the working directory for 24 hours; `resolve_cik()` looks up a ticker's zero-padded 10-digit CIK and official title from that map.
- **XBRL facts** — `fetch_company_facts(cik)` does one bulk request to `data.sec.gov/api/xbrl/companyfacts/` per ticker (rather than one request per line item), which is both faster and friendlier to SEC's rate limits.
- **`CONCEPT_TAXONOMIES`** — a dictionary of ~25 standard line items (Total Revenue, COGS, SG&A, Operating Income, Net Income, EPS, all major balance-sheet and cash-flow lines, etc.), each mapped to an ordered list of the different GAAP XBRL tags real filers use for that concept, since tagging isn't fully standardized across filers.
- **`extract_annual_series()`** — the core normalization logic: buckets each reported fact by the calendar year its period actually *ends in* (not by the filing's own `fy` label, since one 10-K reports several years of comparative figures under the same `fy`), restricts to `10-K`/`10-K/A` annual filings, filters out quarterly/stub periods shorter than 300 days, and deduplicates by keeping the most-recently-filed value for a given year (so later restatements/amendments win).
- **`pull_company_data()`** — orchestrates the above for one ticker across a lookback window, scaling raw dollar values to millions (except EPS, kept as-is, and diluted share count, scaled to millions of shares), and returns a per-concept "coverage" count (how many of the requested years actually had data) alongside the series.
- **Excel model builder** (`build_excel_model`) — writes a formatted `openpyxl` workbook: a title/CIK/source header row, a year header row, then Income Statement / Balance Sheet / Cash Flow / Key Ratios sections. Raw figures are written as plain numbers (in blue, per financial-modeling convention for "hardcoded input"); derived rows (gross profit, margins, YoY growth, total liabilities+equity, the balance check, current ratio, D/E, ROE, ROA) are written as actual Excel formulas referencing other cells by address, including an `=ROUND(Total Assets - (Liabilities+Equity), 0)` balance-check row highlighted with its own fill color.
- **`open_file_in_default_app()`** — best-effort cross-platform "open in Excel/Numbers/etc." (`os.startfile` on Windows, `open` on macOS, `xdg-open` on Linux), with a printed fallback message if it fails.
- **Standalone UI** — defines its own `ipywidgets` panel (ticker box, years box, "Build Model Matrix" button) and displays it at import time, so this module is runnable on its own in a notebook, independent of the rest of the toolkit.

### `filing_extractor.py`

- **`fetch_submissions()` / `get_company_profile()`** — one fetch of SEC's `submissions/CIK*.json` endpoint (reused for both filing history and company profile/SIC code, to avoid hitting it twice per ticker).
- **`get_nth_10k_filing()` / `get_latest_10k_filing()`** — walks the filer's recent-filings arrays for `10-K`/`10-K/A` forms and builds the direct EDGAR document URL for the Nth most recent one (n=0 latest, n=1 the one before, used for year-over-year diffing).
- **`sic_peer_lookup()`** — queries SEC's `browse-edgar` endpoint (Atom/XML output) for other companies sharing a given 4-digit SIC code, parsed with BeautifulSoup's XML parser; returns an empty list (non-fatal) rather than raising if the request fails, since text-mined competitors can still stand on their own.
- **`normalize_filing_text()`** — converts a filing's raw HTML into clean, single-spaced plain text, deliberately flattening rather than preserving paragraph breaks (some filings have hard line-wraps inside a single paragraph that would otherwise be misread as paragraph boundaries).
- **Section extraction** (`extract_business_section`, `extract_risk_factors_section`, and the shared `extract_section()` helper) — finds every occurrence of a heading pattern (e.g. "Item 1. Business") and every occurrence of any of the possible following headings, and picks the start/end pair with the *largest* gap between them. This is what correctly skips the filing's table of contents (where consecutive item headings sit only a few characters apart) and lands on the real section body (which runs thousands of characters).
- **Competitor name mining** (`extract_competitor_candidates`) — two complementary strategies: (1) any capitalized phrase followed by a corporate suffix (Inc., Corp., LLC, plc, etc.) anywhere in the text, which is high-confidence on its own, and (2) any capitalized phrase found within a window around an explicit competition cue word ("compete," "rival," etc.), which catches the much more common case of a competitor named without a corporate suffix at all ("we compete with Apple, Samsung, and Dell"). A stopword list and generic-phrase filter cut down obvious noise.
- **`match_against_ticker_map()`** — the real verification step: cross-references each mined candidate against SEC's actual company titles using word-boundary matching (so a short candidate can't spuriously match inside an unrelated longer word). Confirmed matches get resolved to ticker/CIK; unmatched candidates are kept separately as likely-private competitors (real competitors, just without public financials to compare against) rather than silently dropped.

### `analysis_engine.py`

- **`classify_product_type()`** — scores the Business-section text against a hand-picked list of "consumable/recurring" keywords (subscription, renewal, replenishment, etc.) versus "durable/one-time" keywords (capital equipment, one-time purchase, etc.) and classifies based on whichever scores higher (or "Undetermined"/"Mixed signal" if there's no clear lean). Explicitly flagged in the code as a rough heuristic to be made more robust later.
- **`extract_end_markets()`** — flags which of ~25 common industries (automotive, healthcare, defense, semiconductor, etc.) are simply mentioned by name in the Business section.
- **`diff_risk_factors()`** — splits both years' Risk Factors text into sentences (filtering very short ones) and, for each sentence in the new filing, uses `difflib.get_close_matches` (cutoff 0.6) to check whether anything similar existed in last year's filing; sentences with no close match are flagged as likely new/emerging risk language.
- **`compute_market_share()`** — dependency-injected (takes the ticker-map/CIK-resolution/facts-fetching/series-extraction functions as parameters rather than importing `sec_model_builder` directly, so this module can be tested standalone) function that pulls revenue for a target ticker plus its peer group, and computes each company's share of the combined peer group's revenue per year — explicitly *peer-group* share, not total industry size, since the tool has no way to know true total market size.

### `excel_plugins.py`

- **`build_competitive_landscape_sheet()`** — adds a "Competitive Landscape" tab: a revenue table and a share-of-peer-group-revenue table (sorted by most recent revenue, target company highlighted and labeled), a `Trend` column computed as a live formula (last year's share minus first year's), and a listed-out section for named-but-unmatched (likely private) competitors.
- **`build_industry_factors_sheet()`** — adds an "Industry Factors" tab presenting the product-type classification and its raw keyword-hit scores, the detected end markets, and the list of new risk-factor sentences (word-wrapped, one per row), each section explicitly labeled as a heuristic read to be verified manually.
- **`build_scorecard_sheet()`** — adds a "Research Signal Summary" tab: scores peer-group share trend and YoY revenue growth into a simple +1/0/−1 lean each, sums them into an overall "Bullish/Neutral/Bearish lean," and lists product type and new-risk-count as unscored context. Explicitly labeled on the sheet itself as "informational synthesis... not a recommendation. Not financial advice."

### `competitor_intel_plugin.py`

The orchestrator. `run_full_research_workbook(ticker, lookback)` runs the entire pipeline described in [What it actually does](#what-it-actually-does-end-to-end) — resolve ticker → fetch submissions/profile → pull latest+prior 10-K text → extract Business/Risk sections → mine + confirm competitors (text-mined and SIC-based) → compute market share → classify product type/end markets → diff risk factors → build the base three-statement model → attach the three extra tabs → save. It also defines the notebook's primary UI: a ticker box, a lookback-years box, and a "Build Full Research Workbook" button that runs the pipeline and displays a clickable link to the finished file (plus attempting to auto-open it).

### `RUN_CODE.ipynb`

A single cell: enables IPython's `autoreload` extension (so edits to the `.py` files take effect without restarting the kernel) and imports `competitor_intel_plugin`. Because Python executes a module's top-level code on import, importing `competitor_intel_plugin` transitively imports `sec_model_builder` first — which means **both** UI panels get displayed when you run this notebook: `sec_model_builder`'s simpler "Multi-Concept SEC Financial Modeling System" panel (model only) appears first, followed by `competitor_intel_plugin`'s own "Competitor & Industry Intelligence Plugins" panel (full pipeline). Both are independently functional; this is just a side effect of the import chain rather than a deliberate two-panel design, worth knowing if the extra panel is unexpected.

## Requirements

- **Python 3** with a **Jupyter environment** (Jupyter Notebook, JupyterLab, or VS Code's notebook support) — `RUN_CODE.ipynb` needs a kernel with `ipywidgets` support to render the buttons.
- Third-party packages:
  - [`requests`](https://pypi.org/project/requests/)
  - [`beautifulsoup4`](https://pypi.org/project/beautifulsoup4/) (imported as `bs4`) — note `filing_extractor.sic_peer_lookup()` parses XML with `BeautifulSoup(resp.text, 'xml')`, which needs an XML parser backend (typically **`lxml`**) installed separately; without it, that call will raise a `bs4.FeatureNotFound` error.
  - [`openpyxl`](https://pypi.org/project/openpyxl/)
  - [`ipywidgets`](https://pypi.org/project/ipywidgets/)

There's no `requirements.txt` in the archive. To install:

```bash
pip install requests beautifulsoup4 lxml openpyxl ipywidgets jupyter
jupyter nbextension enable --py widgetsnbextension   # if widgets don't render in classic Notebook
```

## Running it

1. Keep all five files (plus the notebook) in the same folder — the modules import each other by filename, and `RUN_CODE.ipynb`'s own comment notes this ("put all related files in a single folder and run with this file").
2. Open `RUN_CODE.ipynb` in Jupyter and run its one cell.
3. Two widget panels will appear (see the `RUN_CODE.ipynb` note above): a simple ticker/years/button UI for a model-only build, and a fuller ticker/years/button UI for the complete research workbook. Enter a ticker (e.g. `AAPL`, `ILMN`) and click the relevant button.
4. The resulting `.xlsx` is saved to the current working directory as `{TICKER}_three_statement_model.xlsx` or `{TICKER}_research_workbook.xlsx`, and the tool attempts to open it automatically in the OS's default spreadsheet app; a clickable link is also printed in the notebook output regardless.

Each `.py` module can also be read/imported independently for testing (`analysis_engine.py` and parts of `filing_extractor.py` don't depend on the SEC request layer at all, by design, so they can be unit-tested without hitting the network), though only `RUN_CODE.ipynb` gives you the intended UI-driven workflow.

## Practical caveats & things worth knowing

- **A real name and email are hardcoded into the request headers**: `HEADERS = {'User-Agent': 'Sasha Brown s.i.2002@gmail.com'}` in `sec_model_builder.py`. This is exactly what SEC's fair-access policy asks automated tools to provide, so it's correct behavior — but if this code is going into a public repository or portfolio, that's a real personal email address baked into source, and worth swapping for a placeholder (or reading from an environment variable) before publishing.
- **Competitor and product-type detection are text heuristics, not verified facts.** The code's own comments are candid about this: competitor name mining is "a candidate generator, not a verifier," and product-type classification is keyword counting explicitly flagged as needing to become "more robust in later models." Treat both as a starting point for manual review, not a ground truth.
- **"Market share" here means share of the identified peer group's combined revenue, not share of the true total industry.** The tool has no way to know actual total market size, and the sheet's own subtitle says as much — a peer group that's missing a major competitor (e.g. a private company or one the text-mining/SIC lookup missed) will overstate everyone else's share.
- **Risk-factor diffing is fuzzy-sentence-based**, so a risk that's meaningfully reworded (rather than newly added) can still show up as "new," and conversely a genuinely new risk phrased similarly to an old one could be missed.
- **Requires the `lxml` package** for the SIC peer lookup specifically (BeautifulSoup's `'xml'` feature), which is easy to miss since it's not imported directly anywhere — pip installing just `beautifulsoup4` without `lxml` will make that one function fail at runtime, though the pipeline treats that failure as non-fatal (empty peer list) rather than crashing.
- **Local caches are relative to the working directory.** `sec_tickers_cache.json` (and the output `.xlsx` files) are written wherever the notebook happens to be run from — running the notebook from different folders will produce separate ticker caches rather than sharing one.
- **Everything hinges on live network access to SEC EDGAR.** There's no offline/sample-data mode; if SEC's endpoints are unreachable or a specific company has incomplete/late-tagged XBRL data, individual line items will simply come back empty (reflected honestly in the "coverage" counts printed to the notebook) rather than causing a crash — a deliberate and reasonably robust design choice.
- **The Excel outputs are explicitly labeled as non-advice.** The Research Signal Summary sheet states directly on the sheet that it's "informational synthesis... not a recommendation. Not financial advice" — worth keeping that framing if this is shared with anyone else.

## Suggested next steps

- Replace the hardcoded personal `User-Agent` with a configurable value (env var or config file) before any public/portfolio use.
- Add `requirements.txt` (or a `pyproject.toml`) pinning `requests`, `beautifulsoup4`, `lxml`, `openpyxl`, `ipywidgets`.
- Deduplicate the two ipywidgets UI panels that currently both render on import (e.g. only build/display `sec_model_builder`'s standalone panel when that module is actually run directly, via `if __name__ == "__main__"`-style guarding — though notebooks complicate that check since everything runs as `__main__` at the top level).
- Expand `CONSUMABLE_KEYWORDS`/`DURABLE_KEYWORDS`/`COMMON_END_MARKETS` lists, or replace the keyword-count heuristic with a small classifier, as the code's own comments already suggest.
- Add automated tests around the section-extraction and competitor-matching regexes using a handful of saved real 10-K excerpts, since these are the most heuristic-heavy (and most valuable to get right) parts of the pipeline.
- Consider caching fetched company-facts/submissions JSON per CIK (similar to the existing ticker-map cache) to cut down repeat requests when iterating on a single ticker during development.
