# Financial methodology and review questions

These conventions define this model; a computed value is not a claim of economic comparability.

### Definition choices

- ROA and ROE use opening and closing balances (FY−1 and latest FY), with a
  positive average denominator and no negative balance base.
- CAGR uses the first and last **presented** FY and the elapsed year difference,
  including gaps. Missing/invalid first-year figures are not replaced by a later
  convenient base. A single FY cannot produce CAGR. A zero ending value is a
  valid −100% CAGR; the starting value must be positive.
- Net leverage trend is the change in turns between consecutive FYs, latest minus
  previous. It is not a percentage growth rate. Positive change means more leverage.
- Cash conversion is **CFO / EBITDA**, separate from **CCC = DSO + DIO − DPO**.
  Cycle components use averages of FY−1/FY-end balances and 365 days. Revenue
  proxies credit sales in DSO; COGS proxies credit purchases in DPO. These proxies
  and the annual-average approximation can be unsuitable for seasonal businesses.
- ARR uses explicit FY-end `monthly_recurring_revenue` times 12. Supplying MRR
  asserts that the amount is recurring and excludes one-off revenue. ARR is a
  run rate, not recognized annual revenue.
- WACC uses `market_cap` and **separate `market_value_debt`**, not `total_debt`,
  plus explicit CAPM inputs. All weights and assumptions are supplied for the
  fiscal year end; tax is in [0,1], debt cost and ERP are nonnegative. Negative
  risk-free rates and negative beta are permitted. The two-component formulation
  excludes preferred equity. Even with zero debt, this version requires the full
  explicit assumption set rather than inventing unused assumptions.
- EV is the documented simplified `market_cap + total_debt − cash` convention:
  reported debt is a proxy, and preferred equity / noncontrolling interests are
  excluded. EV multiples inherit that convention. P/E needs positive net income.
- FCF is CFO minus nonnegative capex spending; it is not automatically FCFF or FCFE.
  Quick ratio removes inventory only. Negative CFO, FCF, net debt and CCC are allowed
  where meaningful. Invalid signs/bases blank the affected KPI with a reason.

Methodology references: [CFA Institute on CCC](https://blogs.cfainstitute.org/blog/2013/05/21/a-look-at-the-cash-conversion-cycle/),
[OpenStax working-capital examples](https://openstax.org/books/principles-finance-2e/pages/19-1-what-is-working-capital),
and [OpenStax WACC](https://openstax.org/books/principles-finance-2e/pages/17-3-calculating-the-weighted-average-cost-of-capital).
The conventions above specify this application's choices where definitions vary.


## Why is WACC debt different from total debt?

WACC weights represent current opportunity costs of financing and therefore use equity and debt market values. `total_debt` is the reported borrowing balance used in net debt. A separate `market_value_debt` avoids silently treating the carrying amount as market value. For Microsoft FY2025, Note 10 gives about $40.4bn fair value of long-term borrowings including the current portion; that does not supply a marginal borrowing cost, risk-free rate, beta, ERP or marginal tax rate. The model leaves WACC blank until those assumptions are explicitly supplied. This simplified two-component WACC excludes preferred stock and lease financing; document an expanded perimeter before using it for valuation.

## Why does strict Microsoft EBITDA remain unavailable?

The selected statements do not supply a compatible reported EBITDA. CFO's “Depreciation, amortization, and other” contains other adjustments and cannot establish pure D&A. Missing reported EBITDA is an evidence result, not a failed calculation.

An explicit `operating-ebitda-separate-da-v1` recipe adds operating EBIT to **separately disclosed depreciation (Note 6) and intangible amortization (Note 9)**. Its results are approximate reconstructed operating EBITDA, not a reported or management-adjusted measure:

| USD billions | FY2023 | FY2024 | FY2025 |
| --- | ---: | ---: | ---: |
| Operating EBIT | 88.523 | 109.433 | 128.528 |
| Depreciation | 11.0 | 15.2 | 22.0 |
| Intangible amortization | 2.5 | 4.8 | 6.0 |
| Reconstructed operating EBITDA | 102.023 | 129.433 | 156.528 |

The arithmetic retains the disclosed rounding; three decimal places in the sum do not create three-decimal economic precision. No additional finance-lease ROU amortization is added. The recipe runs only when explicitly requested and never replaces a selected reported EBITDA, including a reported zero. Components, source hashes and recipe version are retained in observations, manifest and the demo's Derivations sheet. Dependent ratios inherit the caveat.

Source: [Microsoft 2025 annual report, Notes 6 and 9](https://www.microsoft.com/investor/reports/ar25/). Original archived extracts and line locators are in `research/microsoft_fy2025/raw/`; they are not represented as original downloaded filings.

## Historical equity value and simplified EV

The [Microsoft 2025 proxy](https://www.sec.gov/Archives/edgar/data/789019/000119312525245150/d908201ddef14a.htm), printed pages 53 and 55, identifies the June 30, 2025 closing price as **$497.41**. Multiply by **7,434 million FY-end shares** (annual report Note 15), not weighted-average EPS shares, to obtain **$3,697,745,940,000**. This is approximate historical capitalization because the share count is rounded. The price is not dividend-adjusted close. It is not a current market valuation.

Simplified EV then adds borrowing debt and subtracts cash equivalents. Microsoft short-term investments and finance/operating leases are outside this model's chosen perimeter; preferred equity and noncontrolling interests are not modeled. EV/EBITDA combines both the simplified-EV and reconstructed-EBITDA caveats. P/E uses annual consolidated net income.

## Why two cash conversion metrics?

CFO / EBITDA describes cash conversion relative to operating earnings. CCC estimates days of operating capital tied up: average receivables / revenue × 365 + average inventory / COGS × 365 − average payables / COGS × 365. It is a separate concept. Microsoft's payables include capex, and cost of revenue includes services: its negative modeled CCC is a coarse proxy, not proof of operational superiority.

## Why is ARR still empty?

Company-wide fiscal-year-end recurring MRR is not established by the selected evidence. Cloud revenue, annual recognized revenue, deferred revenue and remaining performance obligations do not become ARR through an alias. The existing ARR formula is explicit MRR × 12. It waits for an appropriate input.

## CAGR and growth periods

YoY requires exactly FY−1. CAGR uses the earliest and latest presented years and their elapsed year difference. FY2023 to FY2025 has two periods. A nonpositive starting denominator does not become a meaningful CAGR through a guessed replacement base. These reported growth measures include acquisitions; they are not organic-growth estimates.
