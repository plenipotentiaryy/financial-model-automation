# Financial Model Automation — Codex Implementation Brief

> **Instruction to Codex:** Build a functioning, testable MVP in the current repository. Do not merely produce another proposal or a mock UI. Inspect the repository first; if it is empty, scaffold a small Python project. Make reasonable implementation decisions without asking questions when the spec already provides a default. Prefer correct, deterministic financial logic over a large feature count. At the end, run the tests, demonstrate the sample end-to-end workflow, and report what works and what remains out of scope.

## 1. Product concept

Build a **deterministic financial model automation tool** that takes financial data for a company, normalizes and validates it, fills a **fixed, reusable Excel financial model**, and produces a traceable, ready-to-review workbook.

The idea is simple:

```text
Financial company data (CSV / JSON / XLSX)
       ↓
Ingestion + column mapping
       ↓
Normalization (units, fiscal periods, currency, definitions)
       ↓
Validation + data provenance
       ↓
Populate predefined Excel Inputs / Assumptions cells
       ↓
Existing Excel formulas calculate KPIs and ratios
       ↓
Completed workbook + validation report + source map
```

**Key principle:** We are NOT building an autonomous AI financial analyst or an LLM wrapper. The formulas, Excel layout, field definitions, and computation rules are predefined. The software automatically supplies the correct inputs, catches data-quality problems, and produces consistent models for different companies.

**Core promise:** *Load company financial data once; obtain a standardized financial model without manually copying values across dozens or hundreds of Excel cells.*

### Target use cases

- Financial analysts creating comparable company profiles.
- Valuation / investment banking / equity research teams using standardized Excel templates.
- Students and hackathon teams automating repetitive financial modeling.
- Eventually, analysis of private European companies with heterogeneous reports.

### What makes this useful

The value proposition is **not** automating arithmetic such as `EBITDA / Revenue`. Excel already does that. The harder, valuable problems are:

1. Mapping many different input field names to consistent financial concepts.
2. Ensuring period, currency, units, accounting scope, and metric definition match.
3. Supplying inputs to a fixed model without breaking existing formulas or formatting.
4. Keeping each number traceable to its source.
5. Identifying missing/ambiguous inputs instead of silently inventing values.
6. Batch-generating comparable models across many companies.

## 2. Existing solutions / competitor context

This is **not a novel product category**. Existing solutions cover overlapping parts of the workflow:

| Product | Relevant capabilities / market positioning | What to learn |
|---|---|---|
| [Daloopa](https://www.daloopa.com/) | Structured public-company financial data connected to Excel models, data updates, source traceability. | Closest conceptual benchmark for data-to-model workflows. |
| [S&P Capital IQ Pro Office](https://www.spglobal.com/market-intelligence/en/solutions/products/resources/sp-capital-iq-pro-office) | Professional financial datasets and Excel integration. | Workflow and analyst expectations; enterprise positioning. |
| [Wisesheets](https://www.wisesheets.io/) | Financial data directly available within Excel / Google Sheets formulas. | Lightweight spreadsheet-oriented approach. |
| [Fathom](https://www.fathomhq.com/) | Financial KPI analysis, reporting, and integrations. | Reporting UX and custom metrics. |
| [Datarails](https://www.datarails.com/) | Excel-centric finance / FP&A automation and consolidation. | Preserving existing Excel workflows. |
| [Microsoft Power Query](https://learn.microsoft.com/en-us/power-query/combine-files-overview) | Importing, cleaning, transforming and combining structured data in Excel. | Baseline functionality we should not claim is innovative. |

The differentiation to investigate **after** MVP: affordable, local-first, transparent handling of private-company data and heterogeneous reports, with strict definitions and source-to-cell audit trails. Do not claim features absent from competitors without new research.

## 3. Scope — first working MVP

Build a **local, offline, CLI-first** Python application. A web UI can come later.

### In scope

- Import standardized financial data from **CSV, JSON, and XLSX**.
- Support **one or multiple companies** through batch processing.
- Map documented column aliases to canonical financial fields.
- Normalize monetary units (`units`, `thousands`, `millions`) and fiscal years.
- Preserve currency and consolidation scope; reject or flag incompatible data.
- Populate a supplied `.xlsx` template using a strict field-to-cell map.
- Provide a **generated example template** if no user template is present.
- Preserve existing template formulas and formatting.
- Produce explicit errors/warnings for missing data, invalid denominators, conflicting periods, etc.
- Produce `source_map` / provenance metadata and a human-readable validation report.
- Include sample data and automated tests.

### Deliberately out of scope for MVP

- LLMs, autonomous agents, prompt-based extraction or external AI APIs.
- Reading arbitrary unstructured PDF financial reports, OCR, web scraping, automatic entity disambiguation.
- Live market-data feeds or paid financial-data integrations.
- Automatically estimating missing figures or converting annual revenue into ARR.
- Automatically guessing risk-free rates, beta, equity risk premiums, or WACC.
- Full DCF/LBO/three-statement forecasting models.
- Account creation, authentication, cloud storage, monetization or a sophisticated UI.

**Important distinction:** Consistent output templates do **not** imply consistently structured input data. Focus on reliable standardized imports first. Parsing inconsistent private-company reports is a separate, significantly harder future module.

## 4. Architecture

Suggested Python 3.11+ stack:

- `openpyxl`: read/write `.xlsx` templates and preserve formulas/styles.
- Standard-library `csv`, `json`, `decimal`, `pathlib`, `logging`.
- `pydantic` (optional but recommended): typed schemas and validation.
- `PyYAML` (or JSON): editable input aliases and template cell mappings.
- `pytest`: unit and integration tests.
- `argparse` or `Typer`: CLI.

Prefer the smallest reliable dependency set. No database is needed for MVP.

Recommended logical modules:

```text
finmodel/
  cli.py                 # validate, build, batch, explain
  schemas.py             # canonical company/period/value records
  ingest.py              # CSV/JSON/XLSX readers
  mapping.py             # aliases -> canonical fields
  normalize.py           # scale, units, currencies, periods
  validate.py            # checks, statuses, issue reports
  template.py            # cell mapping + template protection
  workbook.py            # fill existing workbook + save
  provenance.py          # source-to-field-to-cell records
  outputs.py             # JSON/CSV validation + run manifest
config/
  aliases.yaml
  template_map.yaml
  metric_definitions.yaml   # descriptions, dependencies, validation rules
samples/
  acme.csv
  missing_inputs.csv
  invalid_periods.csv
  template.xlsx             # fixed formulas, generated by init if absent
tests/
README.md
pyproject.toml
```

Module names can change to fit the repository, but keep the concerns separated.

## 5. Canonical input data contract

Start with a **wide-format** file: one company + one fiscal year per row. Sample:

```csv
company_id,company_name,fiscal_year,period_type,scope,currency,monetary_unit,revenue,ebitda,total_debt,cash_and_equivalents
acme,Acme Example,2024,FY,consolidated,PLN,millions,80,,,
acme,Acme Example,2025,FY,consolidated,PLN,millions,100,20,35,5
```

Expected example outcomes, using FY2025 as the latest year:

- Revenue growth = `100 / 80 - 1` = **25%**.
- EBITDA margin = `20 / 100` = **20%**.
- Net debt = `35 - 5` = **30 million PLN**.
- Net debt / EBITDA = `30 / 20` = **1.5x**.
- ARR = **missing inputs** (no recurring-revenue data).
- WACC = **missing assumptions** (no cost-of-capital inputs).

For this example, `total_debt` and `cash_and_equivalents` are year-end balance-sheet values, while `revenue` and `ebitda` are full-fiscal-year flow values. Track that difference rather than treating every metric as measured over the same interval.

### Canonical field candidates

**Required for the sample:**

- `company_id`, `company_name`, `fiscal_year`, `period_type`, `scope`, `currency`, `monetary_unit`
- `revenue`, `ebitda`, `total_debt`, `cash_and_equivalents`

**Optional, to enable additional metrics later:**

- Income statement: `cogs`, `gross_profit`, `ebit`, `net_income`, `interest_expense`, `income_tax_expense`.
- Cash flow: `cash_flow_from_operations`, `capex` (define as a positive spending amount), `dividends_paid`.
- Balance sheet: `total_assets`, `total_equity`, `current_assets`, `current_liabilities`, `inventory`, `accounts_receivable`, `accounts_payable`.
- Market/capital structure: `market_cap`, `share_price`, `shares_outstanding`, `pre_tax_cost_of_debt`, `risk_free_rate`, `equity_beta`, `equity_risk_premium`, `marginal_tax_rate`.
- Recurring revenue: `monthly_recurring_revenue` (MRR) or an explicit recurring-contract calculation dataset.

Each normalized observation must store more than a number. Include, where applicable:

```json
{
  "company_id": "acme",
  "metric": "revenue",
  "fiscal_year": 2025,
  "period_type": "FY",
  "scope": "consolidated",
  "currency": "PLN",
  "canonical_value": "100000000",
  "canonical_unit": "PLN",
  "original_value": "100",
  "original_unit": "millions",
  "source_file": "samples/acme.csv",
  "source_sheet": null,
  "source_row": 3,
  "source_column": "revenue",
  "value_status": "reported"
}
```

Use decimal-safe conversion for monetary amounts where practical; do not round early. The workbook may *display* values in millions while internal normalization uses absolute currency units. Document the display scale clearly.

### Period and compatibility rules

- Do not combine FY, TTM, quarterly, and point-in-time values as though they were interchangeable.
- Preserve the company's fiscal year convention; `2025` means fiscal 2025, not necessarily December 31, 2025.
- Preserve `consolidated` vs `standalone` scope. Do not mix them silently.
- Do not mix PLN, EUR, USD, etc. without an **explicit FX rate, reference date, and conversion policy**. Default: **reject/flag**, not auto-convert.
- Treat `missing`, `not applicable`, `invalid`, `unavailable` as distinct states. Never substitute zero for absent values.
- Detect duplicate conflicting records and surface them instead of arbitrarily selecting one.
- Flag potential interpretation issues such as negative EBITDA where leverage ratios are not meaningful.

## 6. Excel workbook contract

The workbook layout must remain fixed and reusable for any company:

| Sheet | Purpose |
|---|---|
| `Inputs` | Imported financial statement values for the current and previous fiscal years. |
| `Assumptions` | Explicit parameters not necessarily available in statements (tax, cost of capital, etc.). |
| `Calculations` | Predefined Excel formulas. **Not rewritten on each run.** |
| `Report` | Formatted user-facing financial KPI / valuation summary. |
| `Checks` | Missing-data, consistency, and compatibility flags. |
| `Source_Map` | Source file/sheet/row/column for imported values. |
| `Data_Dictionary` | Definitions, units, and required fields for KPIs. |

**Implementation rule:** Fill only cells explicitly whitelisted in `config/template_map.yaml` (typically `Inputs` and specified `Assumptions` cells). Never overwrite formulas, user-provided template design, or unrelated cells. For MVP use a supplied template if available; otherwise create a minimal professional `samples/template.xlsx` with these sheets and working formulas.

Example mapping design (exact addresses may be adapted):

```yaml
metadata:
  company_name: Inputs!B2
  latest_fiscal_year: Inputs!C3
  previous_fiscal_year: Inputs!B3
  currency: Inputs!C4
financials:
  revenue:
    previous: Inputs!B7
    latest: Inputs!C7
  ebitda:
    latest: Inputs!C8
  total_debt:
    latest: Inputs!C9
  cash_and_equivalents:
    latest: Inputs!C10
```

Example fixed formulas inside the generated workbook (conceptual):

```excel
Revenue growth:       =IF(AND(ISNUMBER(Inputs!B7),ISNUMBER(Inputs!C7),Inputs!B7<>0),Inputs!C7/Inputs!B7-1,"")
EBITDA margin:        =IF(AND(ISNUMBER(Inputs!C7),ISNUMBER(Inputs!C8),Inputs!C7<>0),Inputs!C8/Inputs!C7,"")
Net debt:             =IF(COUNT(Inputs!C9:C10)=2,Inputs!C9-Inputs!C10,"")
Net debt / EBITDA:    =IF(AND(ISNUMBER(B4),ISNUMBER(Inputs!C8),Inputs!C8>0),B4/Inputs!C8,"")
```

For the final formula, `B4` means the *Net debt* formula cell on `Calculations`; make the actual cross-cell dependency explicit in the generated template. Missing data must show an empty calculation cell **plus an explicit issue/status in `Checks`**, not an unexplained blank. Prevent incorrect dependencies and circular references.

### Very important: formula recalculation

`openpyxl` **does not calculate Excel formulas**. Writing input cells with `openpyxl` does not automatically produce updated cached formula results. Do not claim otherwise.

- Treat the trusted Excel template as the **single source of truth** for calculation formulas in this MVP.
- Write inputs safely and request full recalculation when opened in Excel (`fullCalcOnLoad`/`forceFullCalc` as supported by the library).
- Detect whether a compatible spreadsheet recalculation engine is installed; optionally support an explicit `--recalculate` mode. Do not assume such an engine is always available.
- If recalculation has **not** occurred, make that limitation visible in the command output / run manifest. Never present stale formula caches as fresh computed values.
- Do not separately hand-maintain a second set of financial formulas in Python. If headless calculation becomes a product requirement later, implement it from **one authoritative, versioned formula registry** and test parity with the Excel output, or require a spreadsheet calculation engine.

## 7. Financial metric definitions and input availability

Design for a catalog of **20–30+ KPIs**, but implement the first small, correct group fully before adding more. Each definition should document:

- Canonical metric ID / readable label.
- Formula and interpretation.
- Required input field IDs and applicable periods.
- Unit (currency, %, x, count).
- Whether the metric applies to every company or only certain business models.
- Data-quality and denominator checks.
- Sources, assumptions and definition choices.

### MVP metrics — implement and test these first

| Metric | Formula / logic | Requirements |
|---|---|---|
| Revenue growth YoY | `Revenue_latest / Revenue_previous - 1` | Comparable FY revenue; nonzero prior revenue. |
| EBITDA margin | `EBITDA / Revenue` | Same FY/scope/currency; nonzero revenue. |
| Net debt | `Total debt - Cash & cash equivalents` | Same reporting date, currency, and scope. |
| Net debt / EBITDA | `Net debt / EBITDA` | Valid annual EBITDA, and positive EBITDA for meaningful ratio. |

### Future / conditional metric catalog

| Metric | Notes / extra requirements |
|---|---|
| Gross margin | Gross profit / revenue; consistent gross-profit definition. |
| EBIT margin | EBIT / revenue. |
| Net profit margin | Net income / revenue. |
| Free cash flow | Defined in the MVP convention as operating cash flow minus **positive capex spending**; disclose definition. |
| FCF margin | FCF / revenue. |
| Current ratio | Current assets / current liabilities. |
| Quick ratio | (Current assets - inventory) / current liabilities, subject to definition. |
| Interest coverage | EBIT / interest expense; validate denominator and sign. |
| ROA | Prefer net income / **average** total assets, requiring prior/current balances. |
| ROE | Prefer net income / **average** equity, requiring prior/current balances. |
| Enterprise value | Equity market value + debt - cash; clarify treatment of minority interest/preferred stock if present. |
| EV / Revenue | Enterprise value / relevant annual revenue. |
| EV / EBITDA | Enterprise value / compatible annual EBITDA. |
| P/E | Market capitalization / annual earnings if meaningful. |
| Revenue CAGR | Requires consistent start/end periods and year count. |
| EBITDA CAGR | Requires consistent start/end periods and meaningful bases. |
| Net leverage trend | Comparable net debt and EBITDA across periods. |
| Cash conversion | Requires an explicit chosen definition. |
| ARR | For recurring-revenue businesses; commonly `MRR × 12`, **not** equal to total annual revenue. |
| WACC | Requires explicit market-value weights and cost-of-capital assumptions. |

**ARR:** Do not infer ARR from revenue unless the data explicitly demonstrates that the revenue is recurring under the selected ARR definition. No MRR or recurring contract data => `MISSING_INPUTS` or `NOT_APPLICABLE`.

**WACC:** Under the simplest two-component formulation:

```text
WACC = [E/(D+E)] × Ke + [D/(D+E)] × Kd × (1−T)
Ke   = Rf + beta × ERP    # CAPM assumption; must be disclosed
```

Here `E` and `D` are appropriate **market values**, `Ke` is cost of equity, `Kd` is pre-tax cost of debt and `T` is the relevant tax rate. Do **not** silently substitute book equity/debt or pull arbitrary internet estimates. Flag insufficient assumptions and limitations explicitly. More complex capital structures require an expanded formula.

Use numeric status codes such as:

```text
OK
MISSING_INPUTS
NOT_APPLICABLE
INCOMPATIBLE_PERIODS
CURRENCY_MISMATCH
SCOPE_MISMATCH
INVALID_DENOMINATOR
AMBIGUOUS_SOURCE
FORMULA_NOT_RECALCULATED
```

## 8. Source-to-cell traceability

Maintain a data lineage record for each populated field:

```text
original source → source locator → canonical field → normalized value → workbook cell → dependent metric(s)
```

Examples:

```text
samples/acme.csv, row 3, column revenue
  → acme / FY2025 / consolidated / revenue / 100,000,000 PLN
  → Inputs!C7
  → Calculations!RevenueGrowth + Calculations!EBITDAMargin
```

Record the actual source location, not a fabricated citation. For `.xlsx` input include the sheet name and source cell. For CSV/JSON input include row/field or JSON path.

Outputs should contain:

- `Source_Map` sheet with source file, locator, canonical field, period, original value/unit, normalized value/unit, destination cell.
- Machine-readable `validation.json` / `run_manifest.json` containing company ID, files, statuses, issues, run outcome and recalculation status.
- Summary printed to CLI: successful inputs, missing inputs, invalid records, output file location.

## 9. Validation and security requirements

Important checks:

1. Required column presence, numeric parsing, normalized units.
2. Company identity, fiscal year and period type.
3. Matching currency and consolidated/standalone scope for dependent calculations.
4. No silent mixing of multiple reporting periods or year-end balance sheet vs yearly flow data.
5. Missing values vs real numeric zero (different!).
6. Division by zero and other undefined ratio cases.
7. Duplicate/contradictory values for the same company/period/field.
8. Preserving formula cells and styling in the original template.
9. No reference to wrong fiscal-year cells when building company models.
10. If a required metric cannot be computed, clearly record *why*.

Security and safety:

- Operate locally by default; don't upload company files externally.
- Only accept `.xlsx` for spreadsheet templates in MVP; do not execute macros.
- Sanitize untrusted **text** written to spreadsheet cells to avoid spreadsheet-formula injection. Do not confuse a legitimate negative numeric amount with a malicious text formula.
- Fail safely on malformed input; never invent a missing financial value.
- Do not overwrite an existing output file without an explicit `--overwrite` option.

## 10. CLI behavior

Implement an interface approximately like:

```bash
# Generates an example workbook, sample CSV, and default mappings
python -m finmodel init

# Preview issues before generating an Excel workbook
python -m finmodel validate --input samples/acme.csv

# Produce a company-specific workbook
python -m finmodel build \
  --input samples/acme.csv \
  --template samples/template.xlsx \
  --output outputs/acme_2025.xlsx

# Process a directory of company files into separate workbooks
python -m finmodel batch \
  --input-dir samples/companies \
  --template samples/template.xlsx \
  --output-dir outputs/

# Explain a metric: formula, required fields, limitations
python -m finmodel explain revenue_growth_yoy
```

You can choose a different, equally clear syntax. Document the actual commands and make them runnable. The sample build should use a clean temporary output folder or create one on demand.

For `batch`, do not mix different companies' input data. Produce distinct workbooks and one aggregate job summary; do not add a comparison workbook until standardization and period comparability are tested.

## 11. Engineering priorities / implementation order

**Phase 1 — functional vertical slice (must complete):**

1. Set up Python project + CLI + tests.
2. Define canonical input schemas and sample CSV/JSON data.
3. Implement standardized CSV reader, optional JSON/XLSX readers.
4. Normalize units, years and metadata; validate missing/duplicate data.
5. Generate an example Excel template with fixed sheets and formulas.
6. Populate only whitelisted input cells without altering calculations/styles.
7. Generate source map, checks and run manifest.
8. Run end-to-end sample and tests; prove expected input values and formula links.

**Phase 2 — usefulness:**

9. Complete CSV, JSON, XLSX parity and configurable alias mapping.
10. Batch processing and clear error recovery.
11. Add more metrics selectively with explicit dependencies.
12. Add formula recalculation integration or transparent status if unavailable.
13. Improve workbook formatting / KPI statuses / documentation.

**Phase 3 — possible differentiation (not MVP):**

14. Strict, rule-based parser adapters for particular financial-report formats.
15. Versioned metric definitions and accounting normalization policies (IFRS/US GAAP adjustments must be explicit).
16. Private-company PDF extraction with source citations and human approval of ambiguous facts.
17. Peer comparison with company-type / period compatibility checks.
18. Optional local or cloud UI; optional licensed market-data connectors.

Do not expand into Phase 3 if the core pipeline is not working and tested.

## 12. Acceptance criteria / definition of done

The MVP is acceptable when all of the following are true:

- [ ] A fresh repository can be set up using documented commands.
- [ ] `init` produces a valid Excel template and sample data.
- [ ] `validate` detects missing and inconsistent inputs.
- [ ] `build` creates a company workbook with **all original formula cells intact**.
- [ ] The generated workbook has `Inputs`, `Assumptions`, `Calculations`, `Report`, `Checks`, `Source_Map`, `Data_Dictionary`.
- [ ] FY2024 revenue 80m PLN + FY2025 revenue 100m PLN, EBITDA 20m, debt 35m and cash 5m yield formulas corresponding to **25% growth**, **20% EBITDA margin**, **30m PLN net debt**, **1.5× leverage** once recalculated.
- [ ] ARR and WACC are **not guessed** when their required inputs are missing.
- [ ] Raw values, normalized values, source locations and destination cells are inspectable.
- [ ] Empty input does not silently become zero.
- [ ] Currency or scope conflicts are explicitly flagged.
- [ ] A zero denominator never silently produces a bogus ratio.
- [ ] Formula recalculation status is reported accurately (including if only recalculation-on-open is available).
- [ ] Multiple companies can be processed into separate workbooks.
- [ ] Automated tests cover the sample plus missing/duplicate/conflicting/invalid scenarios.
- [ ] README explains limitations, supported inputs, conventions, CLI commands and how to open/recalculate in Excel.

### Suggested test cases

1. `test_sample_model_inputs_and_formulas`: correct mapped cells and unchanged formulas.
2. `test_normalize_millions`: `100 million PLN → 100,000,000 PLN` internally.
3. `test_no_missing_to_zero`: blank EBITDA remains missing.
4. `test_currency_mismatch`: PLN/EUR records do not mix.
5. `test_scope_mismatch`: consolidated vs standalone records do not mix.
6. `test_duplicate_conflict`: two different revenues for the same entity/year/field are flagged.
7. `test_denominator_validation`: ratio calculation requires valid denominators.
8. `test_no_arr_guess`: annual revenue alone never becomes ARR.
9. `test_no_wacc_guess`: missing capital market assumptions leave WACC unresolved.
10. `test_formula_preservation`: sample template's formula cells remain intact after export.
11. `test_provenance`: every imported numeric field has a genuine source locator.
12. `test_batch_isolation`: company A's data never appears in company B's workbook.
13. `test_no_stale_cache_claims`: no assertion of recomputed Excel values without a real calculation engine.

## 13. Product decisions and trade-offs

| Decision | MVP choice | Rationale |
|---|---|---|
| AI agent vs script | **Deterministic script** | Consistent formulas, low cost, reproducibility, testing. |
| Spreadsheet vs custom computation | **Existing trusted Excel formulas** | Preserves user methodology and layout. |
| UI | **CLI first** | Validate the core financial data pipeline quickly. |
| Input type | **Structured data first** | Arbitrary statements require a different extraction system. |
| Missing data | **Explicit status** | Avoid fabricated financial metrics. |
| Currency | **No implicit FX conversion** | Comparability and auditability. |
| Financial data storage | **No DB initially** | Sample files, normalized in-memory models and output manifests suffice. |
| Source documentation | **Required** | Analysts must be able to trace numbers to files. |
| Private-company support | **Roadmap** | Potential differentiation, but requires real extraction and definition controls. |

The tool should not promise more automation than it can prove. If the source is ambiguous, the output should say *unresolved / review needed*, not select a plausible number.

## 14. What to report at the end of Codex's implementation

When complete, respond with:

1. What was actually implemented (not planned).
2. The exact runnable commands.
3. Test results and any failed/skipped tests.
4. Paths to a generated sample workbook and validation report.
5. Any limitations, especially whether Excel formulas need recalculation upon opening.
6. The highest-value next development step.

**Final direction:** Build a small but real, reliable *financial data → validated fixed Excel model* pipeline. Favor correctness, auditability, and extensibility over AI features, flashy dashboards, or premature complexity.
