# Prepared-data and Excel reference

## Input contract

One company and fiscal year per row, with required metadata headers:

```csv
company_id,company_name,fiscal_year,period_type,scope,currency,monetary_unit,revenue,ebitda,total_debt,cash_and_equivalents
acme,Acme Example,2024,FY,consolidated,PLN,millions,80,,,
acme,Acme Example,2025,FY,consolidated,PLN,millions,100,20,35,5
```

- **CSV:** UTF-8 (optional BOM), comma-separated, header row, decimal point,
  standard CSV quoting. Ragged rows, duplicate mapped columns and malformed
  quoting fail. Unknown columns are ignored with a warning.
- **JSON:** a nonempty array of row objects using the same fields. Null is
  missing. Decimal number lexemes are preserved without conversion through binary
  floating point. Duplicate object keys fail. Numeric strings are supported.
- **XLSX:** headers on row 1. Automatically select the only nonempty sheet, or
  pass `--sheet Financials`. Multiple nonempty sheets require explicit selection.
  Financial formula cells are rejected; cached source results are not trusted.
- `company_id` is an explicit identifier, not inferred from the name. Conflicting
  names for the same ID fail. XLSX numeric precision may already have been lost
  in the source file; prefer CSV/JSON strings for exact large decimals.
- Fiscal years accept four-digit years or `FY2025` / `FY 2025`. Only `FY` rows
  are supported. The latest presented FY is selected; YoY requires exactly FY−1.
  Missing FY−1 is never replaced by an older available year.
- `scope`: `consolidated` or `standalone`. `currency`: an explicit three-letter
  code. The tool checks its shape and compatibility, not membership in an ISO
  currency registry. No implicit currency conversion occurs.
- `monetary_unit`: `units`, `thousands` or `millions`. Normalize money to absolute
  currency units using `Decimal`; the workbook displays monetary KPI/input
  amounts in millions through number formatting, not by changing stored values.
- Flow fields represent the full FY. Debt/cash and other balance fields represent
  that same fiscal year end. The wide-format contract asserts this timing; the
  importer cannot prove it from financial statements. It never invents a calendar
  year-end date. MRR, if provided, is an explicit monthly run rate at FY end.
- Optional rates are **decimal fractions**, beta is a ratio, shares are counts,
  and share price is currency per share; these fields are not multiplied by the
  row's monetary scale. Capex, when provided, is positive spending.

Documented aliases live in `config/aliases.yaml` (for example `sales` maps to
`revenue`). Matching ignores header case and surrounding whitespace, and
normalizes spaces/hyphens to underscores. There is no fuzzy semantic matching.
`--aliases`, `--template-map` and `--metric-definitions` select alternative YAML
files. Otherwise commands use `config/*.yaml` from the working directory if
present, with packaged defaults when absent. Runtime definitions for all 28
report metrics cannot silently change dependencies or formulas: that requires a
versioned template/code update.

### Missing and invalid values

Empty/null/`missing`, `not applicable`/`n/a`, `unavailable` and `invalid` are
distinct observation states. A numeric zero stays zero. Nonfinite numbers,
grouping separators, percent strings and formulas are rejected as numeric input.

Structural/identity/numeric errors, incompatible period/currency/scope and
contradictory duplicates block the affected company. Equal normalized duplicate
values are merged with **all** original source locators and scales retained.
Missing financial fields are warnings rather than structural errors: a company
can have a partial workbook. Zero denominators and nonpositive EBITDA leave only
the affected KPI blank, with an explicit status. Negative net debt is valid.

## Workbook and calculation behavior

The generated workbook contains `Report`, `Assumptions`, `Calculations`, `Inputs`,
`Checks`, `Source_Map` and `Data_Dictionary`. A supplied trusted `.xlsx` template
must have these sheets and a matching cell map. Formula definitions are created
by `init`; `build` never rewrites financial formulas.

`config/template_map.yaml` is the complete write allowlist. It contains input
cells and bounded report-status, check and source-map regions. Overlaps, formula
destinations, merged-cell intersections, absent sheets and insufficient region
capacity fail. Only allowlisted values are cleared before population; stale
inputs never survive a run. Existing cell styles, formulas and unrelated cells
are preserved. New audit records receive readable formatting. Untrusted text is
stored as literal text in XLSX, not as executable formulas. CSV audit text is
escaped when necessary for safe opening; exact original strings remain in the
JSON manifest and XLSX text cells.

**`openpyxl` does not calculate formulas.** Without `--recalculate`, outputs
request automatic/full recalculation on opening and report
`FORMULA_NOT_RECALCULATED`. Formula caches are not reported as fresh. Open the
workbook in Excel with calculation set to Automatic, use Calculate Now or a full
recalculation, and save it. Opening the workbook later does not update the saved
CLI manifest; rerun with `--recalculate` for a recorded headless calculation.

`--recalculate` requires a locally installed LibreOffice (`soffice` or
`libreoffice`; the standard macOS app location is also detected). It runs a
temporary copy with an isolated profile and a 60-second timeout. All original
scalar formula cells must survive with equivalent references and non-error
results. Only verified result caches and cache types are transplanted into the
original workbook ZIP; the LibreOffice-resaved design is discarded. The original
formula XML, styles and other ZIP parts remain in the exported book. Empty
numeric caches are not accepted as recalculated values. A subsequent openpyxl
save would discard these caches; the application does not do that.

Unavailable engines, timeouts, formula changes or spreadsheet errors fail an
explicit recalculation request, produce a failure report, and do not publish a
new workbook. No duplicate KPI engine is maintained in Python; optional source reconstruction recipes are separate from KPI formulas. Input-readiness checks only inspect prerequisites and
denominators.

Report/check statuses describe the **import at build time**. Financial formulas
react to input edits in Excel, but static audit statuses, provenance and manifests
do not: rerun the pipeline after changing source data. Recalculation-on-open stays
enabled even for a successfully recalculated workbook.

## Outputs, batch and failure handling

For `outputs/all_metrics_2025.xlsx`, the adjacent `outputs/all_metrics_2025.reports/` contains:

- `validation.json`: issues, KPI readiness and calculation state.
- `validation.txt`: readable explanations.
- `source_map.csv`: source locators, original/normalized values, destinations,
  measurement bases and dependent metric IDs. Optional/unselected observations
  are retained with no destination instead of pretending they were imported.
- `run_manifest.json`: all normalized observations and original sources, selected
  company/FY, input/template SHA-256 hashes, effective configuration hashes,
  timestamp, application version, outcome, counts and recalculation status.

`validate` creates per-company reports plus `validation_summary.json` without
creating a workbook. `batch` reads supported files directly inside its input
directory, merges records by company ID, and creates one workbook/report set per
company plus `batch_summary.json`. File names contain a sanitized ID, hash suffix
and latest fiscal year. Bad companies/files do not stop independent good ones;
the aggregate command still exits with failure. Conflicts across files are
checked. Use a separate output directory to avoid ingesting prior exports.

For multi-company input with `build`, select `--company <exact_id>` or use batch.
Reports are generated for identified companies with blocking data errors. When
an import has no usable identity, `build` writes reports with `company_id: null`;
`validate`/`batch` preserve the file-level failure in their aggregate JSON report.

Exit codes: `0` successful processing (warnings may remain); `1` invalid data,
export/recalculation failure or refused replacement; `2` CLI argument errors.
Warnings such as missing ARR/WACC are expected for the minimal Acme fixture;
the full all-metrics fixture supplies all required inputs.

Outputs are staged before publication. Existing output files and report sets
require `--overwrite`; source files and the template cannot be output targets.
Report directories containing unrelated files are never deleted. On a failed
overwrite run, a previous workbook remains untouched; the new failure manifest
has `output_file: null`, so its existence is not evidence of a successful run.

## Tests and supported boundaries

```bash
python -m pytest -q
python -m pytest -q -m 'not engine'
```

Tests cover format parity, precise normalization, availability states, aliases,
period/currency/scope errors, duplicate conflicts, provenance, formula/style
preservation, text injection, batch isolation and partial failure, output
protection, stale caches and real LibreOffice recalculation. Engine tests skip
with an explicit reason when LibreOffice is unavailable. A real-engine test
changes revenue and verifies growth and margin update through the Excel formulas.
Full-catalog tests verify every result in both Calculations and Report across
CSV/JSON/XLSX, each metric's missing dependencies, and parity between validation
and Excel guards for zero denominators, invalid assumptions, signs and periods.

Only ordinary `.xlsx` models and scalar formulas are supported. No VBA execution,
external workbook links, array/shared formula recalculation, or guarantee of
round-tripping arbitrary advanced Excel extensions is provided. Use trusted
templates; test your specific template's supported features before adopting it.
Inputs exceeding Excel's 15-significant-digit precision/range are rejected on
export instead of silently rounded. Full exact normalized values remain in the
failure report.

The prepared-data contract above remains strict. The new candidate workflow retains unknown fields, accepts explicitly checked accounting-number formatting, and supports text PDFs. See [review workflow](review.md). OCR, automatic FX, desktop installers, local OpenJev weights, conversational assistants and forecasting remain outside v0.3.
