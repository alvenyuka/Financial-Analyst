# Financial-Analyst: methodology

The detailed write-up behind the [README](../README.md): how the Apple model is built, how it is checked, what
an independent review found, and what is still out of scope. Every figure here is printed by
`validate_model.py` or read from the workbook.

## Contents

1. [Model structure](#model-structure)
2. [Forecast](#forecast)
3. [Valuation](#valuation)
4. [Checking the model](#checking-the-model)
5. [Review history](#review-history)
6. [Limitations](#limitations)
7. [Roadmap](#roadmap)

## Model structure

`Apple/Apple_Financial_Model.xlsx` has one input sheet and every other sheet calculates from it.

| Sheet | Contents |
|---|---|
| Cover | Scope, sources, colour conventions, preparer |
| Dashboard | Headline metrics and charts, all linked to the statements |
| Assumptions | Every driver: segment growth and margins, operating expenses, working-capital days, balance-sheet drivers, WACC inputs, valuation date, share count, reference price, football-field multiples, and the Bear / Base / Bull switch in `B62` |
| IS / BS / CFS | Statements: FY2021 to FY2025 actuals and FY2026E to FY2029E forecast |
| Ratios | Profitability, returns, leverage and working-capital ratios |
| DCF | Free cash flow build, discounting, valuation summary, two sensitivity grids, tornado and football field |
| Validation | Each historical line against its 10-K page, structural checks and forecast sanity bands |
| Pivots | Pivot summaries of the statements |

Inputs are blue, calculations black, cross-sheet links green. There are no macros, external links or add-ins.

## Forecast

- **Revenue** by product line (iPhone, Mac, iPad, Wearables) and Services, each with a scenario growth rate;
  geography follows the same total.
- **Margins and costs**: gross margin by segment, operating-expense growth, a scenario tax rate (Base 15%: the
  FY21 to FY25 average excluding FY24's one-off EU State Aid charge).
- **Working capital** from days (receivables, inventory, payables) and percentages of revenue (vendor
  receivables, other current and non-current assets and liabilities, deferred revenue), each based on the
  FY24 to FY25 average.
- **Capital and financing**: CapEx and D&A as a share of revenue, scheduled debt maturities by scenario,
  commercial paper and short-term securities held at the FY25 balance, and capital returns as a share of free
  cash flow.
- **Cash** is the result of the cash flow statement. A minimum-cash rule sells long-term marketable securities
  when forecast cash would fall below Apple's lowest FY21 to FY25 year-end cash ($23.6bn). The balance sheet
  therefore balances because the cash flow is complete, not because a line absorbs the difference.

## Valuation

- **Valuation date**: 28 Sep 2026, the date of the $338.40 reference close (`Assumptions!B60`).
- **Free cash flow**: EBIT × (1 − scenario tax rate) + D&A − CapEx − change in working capital.
- **Discounting**: each year's cash flow is discounted by the years from the valuation date to its fiscal
  year-end. FY2026 ended on 26 Sep 2026, before the valuation date, so its cash flow is not discounted; its cash
  is in FY26E net cash instead. FY27E to FY29E are discounted 0.99, 2.01 and 3.00 years.
- **Terminal value**: Gordon growth at 2.5% on FY29E cash flow, discounted from the FY29E year-end. It is 83%
  of enterprise value.
- **WACC 9.34%**: cost of equity 10.0% (CAPM, beta 1.1), after-tax cost of debt 3.4%, 90/10 weights.
- **Result (Base)**: enterprise value $2.05tn, plus FY26E net cash $41.4bn, equity value $2.09tn, $139.50 a share.
  Bear and Bull, recorded with each scenario active: $75.86 and $265.54.
- **Reverse DCF**: the validator solves for the discount rate at which the same cash flows give $338.40:
  5.30%.

## Checking the model

The Validation tab reports "ALL CHECKS PASS", but that tick is computed by the same spreadsheet whose
correctness is in question. `validate_model.py` recomputes everything from the raw cells instead:

| Family | Result on the Apple model |
|---|---|
| Accounting identities (income statement down the page, both sides of the balance sheet, the equity split, cash flow sections and roll-forward, cross-statement ties) | 20 of 20 |
| Validation-tab lines traced to the IS, BS or CFS cell they were transcribed from, five years each | 56 of 56 |
| Valuation figures: working capital rebuilt from the balance sheet, the forecast balance sheet re-added from its lines, forecast cash against the cash flow statement, WACC, free cash flow, discount periods, present values, terminal value, enterprise and equity value, net cash, price, upside, sensitivity-grid centre and corners, discounted multiples | 23 of 23 |

One advisory flag is explained rather than suppressed: in FY21 to FY23 Apple's cash flow statement reconciled
to cash, cash equivalents and restricted cash, while the balance-sheet line excludes restricted cash, so the
two differ by 989, 1,331 and 772 million. From FY24 they agree.

The validator is itself tested. `tests/` copies the real model, breaks one thing, and asserts that thing is
caught: a total that no longer adds up, a duplicate row label, mismatched year headers, a statement that no
longer matches its transcription, a mis-discounted terminal value, a wrong price, undiscounted multiples, a cash
flow that reads the wrong balance-sheet year, a forecast balance sheet that does not balance, and net cash from
the wrong year. CI runs those 17 tests first and the real model second.

## Review history

An independent spreadsheet review in October 2026 found four errors that the workbook's own checks and the
earlier validator both missed, because the validator reused two of the workbook's cell references:

1. The cash flow statement read each forecast year's working-capital and debt movements from the year before,
   and skipped FY28 entirely.
2. "Net cash (FY25)" was FY24 net cash.
3. Nine FY2025 balance-sheet lines did not match the 10-K, although their totals did; for example current term
   debt was 0 against 12,350 filed.
4. Long-term securities balanced the forecast balance sheet as a plug, so the balance check could not fail, and
   two forecast inputs were negative balances.

All four are fixed: the cash flow reads the right year, net cash comes from FY26E (the last balance sheet
before the valuation date), the Validation tab now carries each balance-sheet line separately with its filing
citation, and the plug is replaced by stated drivers and the minimum-cash rule. The validator now rebuilds
working capital, net cash and the forecast balance sheet from the statements, and has tests for each of these
faults. The valuation moved from $135.98 to $139.50.

The review also removed a second, hand-typed model hidden on the Dashboard (its charts now read the income
statement), consolidated two scenario switches into one, corrected the cover page, moved the NOPAT tax rate to
the scenario rate, and deleted 48 unused named ranges and their documentation sheet. An earlier version of the
workbook (before September 2026) had reported $240.05 on its Bull scenario against an undated $232.50, with the
terminal value discounted one period too many.

## Limitations

- **Forecast drivers are assumptions.** The sanity bands on forecasts are wide (operating margin 25% to 40%),
  so passing them is a weak test.
- **Sensitivity tools are simplified.** The growth × margin grid and the tornado use a one-stage build, so
  their centres differ from the headline price; the WACC × growth grid is the one tied to the main DCF.
- **Bear and Bull are recorded values.** Switching `Assumptions!B62` recalculates them; the validator prints
  them as recorded.
- **Buybacks retire shares at the reference price** in every forecast year, a simplification that affects EPS
  but not the DCF.
- **The FY2026 first-quarter statements** are in `Source_filings/` but not yet used.

## Roadmap

- [x] Apple three-statement model and DCF, every historical line traced to its filing
- [x] Independent validator for the statements and the valuation, with fault-injection tests
- [x] Review fixes: cash flow timing, net cash year, FY2025 lines, no balancing plug, valuation date
- [x] Extend the validator to the Ratios tab (19 rows, rebuilt from the traced line items)
- [ ] Safaricom (SCOM), Nairobi Securities Exchange
- [ ] Equity Group Holdings (EQTY), Nairobi Securities Exchange
