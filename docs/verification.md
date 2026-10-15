# v0.3 release verification

Checked on 2026-10-15 with Python 3.12.13 and a real local LibreOffice installation.

| Check | Result |
| --- | --- |
| Full test suite in a clean locked environment | 211 passed, 0 skipped |
| Test dependency diagnostics | One Starlette/AnyIO deprecation warning; no failing test |
| Existing financial test base | All 165 retained tests pass |
| Candidate/derivation additions | 46 tests pass; Jev responses are explicitly offline fixtures |
| Clean installation | `UV_PROJECT_ENVIRONMENT=outputs/clean_env uv sync --python 3.12 --locked --extra dev` succeeds |
| CLI demonstration | `init → validate → build → batch`, plus `explain`, complete |
| Full fictional fixture | 28/28 ready; real Excel caches match the independent oracle |
| Candidate CLI | Extract, missing-key verify, local review and approved-only build exercised |
| Native browser check | Safari: source evidence visible, two manual acceptances saved, coverage changes 0 → 1/28 |
| Review-built workbook | Real recalculation returns EBITDA margin 20% from reviewed 20/100 inputs |
| Restart / reject / correction / conflict / CSRF | Automated API/SQLite tests pass |
| Seven core worksheets | Rendered and visually inspected; extra source/derivation sheets retained |
| Microsoft strict | 17 computed, 11 explicitly unavailable, all 28 verified |
| Microsoft reconstructed EBITDA | 22 computed, 6 unavailable, all 28 verified |
| Microsoft reconstructed EBITDA + historical cap | 26 computed, 2 unavailable, all 28 verified |
| Microsoft formula preservation | All original formula strings retained in all three variants |
| Edited-input check | Change revenue in a separate workbook and recalculate; Report growth and margin update |
| Downloadable demo books | Byte-identical to the audited recalculated outputs; SHA-256 in the case verification |
| Publication hygiene | No API keys, `.env`, SQLite runs or local runtime environments included; original instruction MD kept locally and ignored |

Reproduce the tests with `python -m pytest -q`, and the 84 Microsoft metric checks with `python research/microsoft_fy2025/replay_v03.py --overwrite`. The latter also refreshes the three downloadable books and their hashes. Full evidence: [Microsoft verification](../research/microsoft_fy2025/v03/verification.json).

## Live Jev boundary

**NOT_RUN.** `TYPESAFE_API_KEY` was not configured. There were zero network requests to TypeSafe, zero live tokens and zero live charges in this release check. No live auto-accept error rate, manual-review fraction or calibrated accuracy can be inferred. These measurements remain null / not measured. Offline tests cover high scores, abstention, incomplete responses, wrong model version, retries, budget refusal and source tampering, but do not measure semantic accuracy.

To run a live smoke test, configure the key locally, extract a small public source set, then run `finmodel verify --run <run-dir> --provider jev --max-cost-usd 0.05`. The local verification JSON records responding model, per-check probabilities, usage, actual/unknown-reserved costs and review fraction. Compare admissions against independently labelled evidence before claiming any error rate.

## Source acquisition boundary

Microsoft annual-report evidence and the SEC proxy closing price were verified against official pages. Direct downloads of the original annual-report HTML and proxy PDF returned HTTP 403; archived extracted evidence is clearly labelled. The generic PDF parser is exercised on real PDF files generated as labelled test fixtures, not claimed as a successful direct SEC PDF download. OCR, local OpenJev and a conversational assistant are not implemented.
