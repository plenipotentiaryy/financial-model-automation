# Candidate review workflow (v0.3)

## Two explicit trust boundaries

`build --input` imports prepared company/FY tables under the original strict contract. `build --run` imports only current approved candidate revisions from a local SQLite review run. There is no fallback from the second mode to unverified prepared data.

Extraction is conservative and deterministic: documented aliases propose a field; unknown labels stay in the queue. Jev evaluates a proposal, never invents numbers. A person can correct an unknown label such as monetary `Gross margin` to `gross_profit` after inspecting the evidence. Unknown abbreviations such as `ORRR` are not automatically treated as ARR.

## Inputs and coordinates

```bash
finmodel extract --input report.csv report.json report.xlsx report.pdf --run-dir outputs/my_run
finmodel extract --input statements.xlsx --sheet Financials --run-dir outputs/my_run
finmodel extract --input report.pdf --context company_context.json --run-dir outputs/pdf_run
```

`company_context.json` supplies claims for later checking, for example:

```json
{"company_id":"example","company_name":"Example Corp","period_type":"FY","scope":"consolidated","currency":"USD","monetary_unit":"millions"}
```

Sources must contain the actual context for automatic Jev admission; a supplied company name is not evidence by itself. `fiscal_year` can be supplied when no adjacent year headers are available. The parser recognizes adjacent year columns in row-label tables and PDF text words. PDF candidates retain the complete page text, one-based page and point-coordinate bounding box. Table candidates retain row/column, sheet/cell or JSON path; CSV also records physical line spans. An image-only page yields `OCR_REQUIRED`. No OCR is performed.

Accounting parentheses, minus signs and unambiguous thousands separators are normalized with Decimal. Candidate monetary scales include units/thousands/millions/billions; the prepared-table contract remains unchanged. Percent strings may only populate rate fields. Monetary `Gross margin` stays unknown until its meaning is established. Mixed D&A-and-other and ROU lines are not promoted to clean D&A.

PDF layout extraction is not a universal statement parser. Footnotes, multi-line labels, crossed year headers, split currency symbols and prose can generate incomplete proposals. They must pass semantic checks or be corrected/rejected manually. No context is silently truncated to fit the model. Source files are referenced in place; keep them available and unchanged. Re-extracting changed bytes creates new candidates; old hash-bound approvals become invalid.

## Versioned records and persistence

`<run-dir>/.finmodel/review.sqlite3` stores Candidate v1, VerificationResult v1 and ReviewDecision v1 records, immutable candidate revisions, raw provider answers, request accounting and an append-only event history. Repeating extraction of the same file/position is idempotent; identical values from different files retain separate provenance. Provider caches include the full evidence/proposal, policy, threshold and pinned model version. New revisions cannot reuse an earlier approval.

`.finmodel/`, generated outputs and `.env*` are ignored by Git. API keys are read only from `TYPESAFE_API_KEY`, never written to reports. `verify` sends candidate source excerpts and metadata to TypeSafe; extraction, review and Excel build remain local. Choose files accordingly. The adapter does not download reports or upload unrelated workspace content.

## Provider contract and cost boundary

The [official TypeSafe API](https://docs.typesafe.ai/api) accepts a `state`, typed `questions` and model. Six choice questions independently assess:

1. Financial meaning of the proposed metric.
2. Binding between the number and source row/location.
3. Fiscal period.
4. Currency, scale and dimension.
5. Company and consolidation perimeter.
6. Context sufficiency.

Each offers `match`, `other_metric` and `insufficient_context`. The chosen result, full distribution, returned model and token usage are validated. Auto-admission requires **every** match probability ≥0.99, exact model `jev-1.13.0`, and no deterministic errors or unresolved conflicts. The threshold is experimental; model probabilities are not calibrated correctness guarantees. Source text is untrusted data, not instructions.

Python checks source hash, identity, numeric parsing, scale/dimension, FY, currency/scope compatibility and conflicting duplicates. A response cannot override these. API failure, malformed/incomplete output, version mismatch, excessive context or budget exhaustion leaves candidates in the queue. Missing key gives `NOT_RUN` and exit 1. Low confidence is a normal review outcome, not a transport error. Processing errors return exit 1.

[Model pricing](https://docs.typesafe.ai/models), checked 2026-10-15: pinned Jev input price $0.042/million tokens, output free. Pricing can change; update the policy before relying on another tariff. The budget is **cumulative for the run**, not reset by restarting the command. Each attempt reserves a conservative full 64,000-token input allowance ($0.002688); reported valid usage replaces that reservation. A failed/unknown-cost request keeps the reservation. Therefore a budget below one reservation makes no API call. Retries are capped at three attempts, limited to transient/network errors; bounded backoff never exceeds five seconds per wait. Cached answers incur no new request cost. Oversized payloads (>24,000 UTF-8 bytes) abstain conservatively rather than truncating.

The run's `.finmodel/verification.json` retains results and request accounting, actual usage and reserved unknown costs. `false_auto_accepts` is null and accuracy is `NOT_MEASURED` without a labelled live evaluation. No synthetic result is represented as live evidence.

## Local review and export

```bash
finmodel review --run outputs/my_run --port 8765
# For a manually opened browser: add --no-open
finmodel build --run outputs/my_run --company example --output outputs/example.xlsx --recalculate
```

Only `127.0.0.1` is bound. The service checks Host, exact Origin and a per-process session token on modifying requests; tokens change on restart. It does not expose arbitrary filesystem paths. Candidate snippets and labels use text rendering, not HTML interpolation. Source downloads are limited to recorded candidates whose current hash matches.

Accept/reject require a reason. Correcting mapping or metadata creates a fresh revision; correcting the raw number requires correcting/re-extracting the source. Explicit selection atomically rejects conflicting alternatives. Source corruption or missing identity/period/unit/number cannot be overridden with Accept. Decisions survive a server restart. The page shows source context, coordinates, checks/history, conflict IDs and approved-input KPI coverage. Unresolved competitors can invalidate an earlier approval on export.

The output's `review_audit.json`, manifest and Source_Map tie imported observations to candidate IDs/revisions/source hashes. Manifest `input_mode` distinguishes `approved_candidates` from `prepared_data`. All 28 KPI formulas remain in Excel. Static coverage, audit and provenance describe build time; editing Excel inputs does not rewrite review history.

## Optional source reconstruction

The explicit prepared-data command supports `--derive-ebitda component_evidence.json`. Use `research/microsoft_fy2025/v03/ebitda_components.json` as the versioned example (paths relative to repository root). Identity/currency/perimeter, component hashes and years must match; selected reported EBITDA is preserved. `--derive-ebitda` is refused with `--run` so external component files cannot bypass candidate approval. The Microsoft replay is a separately documented, reviewed source adapter; it does not claim Jev approval.

Local OpenJev models and a conversational assistant can later implement the provider interface. They are not bundled or advertised as implemented here.
