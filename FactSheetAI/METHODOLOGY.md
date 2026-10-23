# Methodology and current evidence

## Research question

The intended question is how much equity exposure overlaps when an investor holds four large-cap funds. This release does not answer it for the requested four-fund set: the ICICI Prudential Large Cap monthly PDFs do not contain complete stock ISINs, and some small equity positions are aggregated. The app refuses to substitute a name match or another scheme. Results currently compare only pairs with available ISIN-level snapshots.

## Identity and source selection

AMFI scheme codes identify plans. `config/schemes.json` maps the user-selected Direct Growth schemes and peers to exact codes taken from the archived AMFI NAV file. Stock ISIN from the monthly AMC disclosure is the join key. No ticker or fuzzy-name join is performed. Original AMC bytes remain in `data/raw/amc/`; source links, month, role and hashes are in the source index and generated manifest.

The AMFI company list is effective-dated. Each portfolio month uses the latest list whose effective date precedes or equals that month: 30 June 2025 for Jul–Dec 2025, 31 December 2025 for Jan–Jun 2026. A later list is never applied retrospectively. Three archived half-yearly workbooks are retained; the July 2026 list is a future effective window relative to the observed portfolio sample.

## Parsing and controls

HDFC, Nippon India, SBI and Tata workbook adapters select the configured fund from a named index/scheme sheet, extract rows within the equity section, preserve source-row spans, attach a SHA-256 and parsed timestamp, and reject ambiguous fund matches or conflicting duplicate security names. Duplicate ISIN weights are summed only when names agree after source footnote markers are removed. SBI's April 2026 workbook uses `#` for one weight: its percentage is derived from that same row's disclosed market value and the workbook's `GRAND TOTAL (AUM)`, with the exact derivation recorded in the provenance span. Other unsupported weights are errors. The ICICI factsheet PDFs are preserved but marked `factsheet_reference_only`, excluded from overlap.

The normalized fallback CSV must describe one scheme-month. Pydantic validates holding records and weight bounds. This is not OCR, and parser confidence values are adapter declarations rather than a measured probability. There is no labelled gold set; precision and recall are therefore unknown. NAV/AUM reconciliation is incomplete and should be treated as a release blocker for accounting-grade use.

## Calculations

For scheme A and B, let `wAᵢ` and `wBᵢ` be equity weight in percentage points of NAV for ISIN `i`:

`Overlap(A,B) = Σᵢ min(wAᵢ,wBᵢ)` over the intersection of their ISIN sets.

This treats the shared portion of each holding as duplicated exposure. It does not add weights for distinct securities and does not estimate statistical correlation. Jaccard is `|A∩B|/|A∪B|` and ignores weights. The pairwise matrix is symmetric; its diagonal is each portfolio's equity weight, capped at 100 percentage points.

Top-10 concentration sums the ten largest equity NAV weights. HHI is the sum of squared fractional equity weights (`weight_pct / 100`); `1/HHI` is the effective stock count. Style exposure sums equity weights whose ISIN appears in the effective AMFI top-100 set; below 80% triggers a mandate alert. Monthly turnover is half the sum of absolute changes over the union of prior/current ISINs. Active share is half the absolute difference across constituents and is only meaningful when the chosen benchmark holdings are valid and comparable.

The pandas merge and pure-Python implementation independently calculate pair overlap; the response includes whether both paths match. Property-based tests exercise analytic invariants. No LLM contributes a number or conclusion.

## Findings and scope limits

There are no four-requested-fund findings in this release. Valid outputs are limited to the four scheme codes with full monthly source files (HDFC, SBI, Nippon India, Tata) plus the HDFC Nifty 50 control, and to months where the correct AMFI list exists. The repository's demo compares selected available funds; this is a technical validation, not a conclusion about an investor's four-fund portfolio. Missing scheme-months are listed in `data/missing-source-files.json`.

Weights are source-reported percentages and may not sum to 100 because cash, debt, derivatives, disclosure conventions or parser scope can differ. The app only includes rows identified as equity. Do not interpret this output as full NAV reconciliation, benchmark attribution, forecast, suitability assessment, or investment advice.
