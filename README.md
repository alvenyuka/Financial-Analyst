# Financial-Analyst

Financial models of listed companies, built from their SEC filings and checked line by line by an independent
Python program. The Apple model values the shares at **$139.50 against a market price of $338.40**: to justify
that price, investors must accept about a **5.3% return** on Apple's cash flows, against the 9.3% a standard
cost-of-capital estimate gives. Every historical figure, accounting check and valuation figure reconciles.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![models validated](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml/badge.svg)](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml)

![Apple valuation range by method: DCF $123 to $162 against a share price of $338.40](images/apple_football_field.png)

## Contents

1. [Business problem](#business-problem)
2. [Dataset](#dataset)
3. [Methodology](#methodology)
4. [Results](#results)
5. [Business impact](#business-impact)
6. [Key insights](#key-insights)
7. [Limitations](#limitations)
8. [Repository structure](#repository-structure)
9. [How to run](#how-to-run)
10. [Documentation](#documentation)
11. [License](#license)

## Business problem

Investment, credit and budgeting decisions rest on spreadsheet models, and a spreadsheet's own "all checks
pass" cell is computed by the same formulas it is meant to check. A reviewer needs to know three things before
relying on a valuation: that the historical figures match the filings, that the statements are internally
consistent, and what the market price implies about the assumptions.

This repository builds models to be audited. Every historical figure cites the filing and page it came from,
every input sits on one assumptions sheet, and a separate program recomputes the arithmetic from the raw cells.
The first complete model covers Apple Inc.

## Dataset

| Source | Coverage |
|---|---|
| Apple 10-K filings (SEC EDGAR, CIK 0000320193) | Fiscal years 2021 to 2025: income statement, balance sheet, cash flow statement |
| Apple investor relations | Segment revenue and gross margin (Products, Services) |
| Market data | Close of $338.40 on 28 Sep 2026; 52-week range $245.27 to $341.07 |
| Analyst targets | 44 analysts, low $215, median $340, high $405 (stockanalysis.com, 29 Sep 2026) |

The filings for fiscal 2023 onward are in `Apple/Source_filings/`; fiscal 2021 and 2022 cite the fiscal 2022 and
2023 10-Ks. The forecast runs to fiscal 2029.

![Apple income statement, FY2021 actuals to FY2029 forecast, with inputs in blue](images/apple_income_statement.png)

## Methodology

```mermaid
flowchart LR
    A[SEC 10-K filings] --> B[IS, BS, CFS FY21-FY25]
    B --> C[Scenario drivers: Bear, Base, Bull]
    C --> D[Forecast FY26-FY29]
    D --> E[Unlevered FCF, WACC]
    E --> F[DCF, reverse DCF, football field]
    B --> G[validate_model.py]
    F --> G
```

1. **Historicals** transcribed from the 10-K filings and press releases, each line cited to a page.
2. **Forecast** from segment revenue growth and gross margins, operating-expense growth, working-capital days
   and capital returns, with no balancing plug: cash comes from the cash flow statement. One cell switches every
   driver between Bear, Base and Bull (`CHOOSE(MATCH(...))`).
3. **Valuation.** Unlevered free cash flow discounted at a CAPM-based WACC (cost of equity 10.0% at beta 1.1,
   after-tax cost of debt 3.4%, 90/10 weights) from the 28 Sep 2026 valuation date to each fiscal year-end, a
   Gordon terminal value at the FY2029 year-end, two sensitivity grids and a football field comparing methods.
   FY2026 ended two days before the valuation date, so its cash is counted in net cash rather than discounted.
4. **Reverse DCF.** The validator solves for the discount rate at which the same cash flows justify the market
   price.
5. **Independent validation.** `validate_model.py` re-derives the statement identities, traces each transcribed
   line to its source cell, rebuilds working capital from the balance sheet and rebuilds the DCF; its tests break
   a copy of the model to prove each fault is caught.

## Results

| Apple, Base scenario | Value |
|---|---:|
| Accounting identities reconciled | 20 of 20 |
| Transcribed line items traced to source cells | 56 of 56 |
| Valuation figures rebuilt independently | 23 of 23 |
| Enterprise value | $2.05tn |
| DCF implied share price | **$139.50** |
| Bear / Bull implied price | $75.86 / $265.54 |
| Terminal value share of enterprise value | 83% |

![Apple DCF: free cash flow build, discounting and valuation summary](images/apple_dcf.png)

## Business impact

What the model tells an investment or credit committee about Apple at $338.40 (all figures printed by
`validate_model.py`):

| Measure | Value |
|---|---:|
| Equity value, Base DCF | $2.09tn |
| Market value at $338.40 (15.0bn diluted shares) | $5.08tn |
| Premium the market pays over the Base case | **$2.98tn** |
| Discount rate the market price implies | **5.30%**, against 9.34% |
| Bull-case price against market | $265.54, 22% below |

At a 9.34% cost of capital the market price cannot be reached even in the Bull case. It is consistent with
investors accepting about a 5.3% return on Apple's cash flows, or with growth well above every scenario here.
A committee can use the model to see which assumption carries the price, rather than argue about the output.

## Key insights

- **The discount rate carries the valuation.** The terminal value is 83% of enterprise value, so a half-point
  change in WACC moves the price more than any forecast-year assumption.
- **Services margins drive the forecast.** Products gross margin is held at its fiscal 2024-2025 average of
  37%, while Services sits at 76% on a trend rising about 1.4 points a year.
- **Spreadsheet checks need an outside check.** An independent review found the cash flow statement reading each
  forecast year's working-capital movement from the year before, and net cash taken from the wrong year, while
  the workbook's own checks showed all green. Both are fixed and the validator now rebuilds those figures from
  the balance sheet itself; its tests corrupt one cell at a time to confirm each break is reported. The
  [methodology](docs/METHODOLOGY.md#review-history) lists every finding.

![Validation tab: each historical income-statement line for FY21 to FY25 ticked against the 10-K page it came from](images/apple_validation.png)

## Limitations

- **Bear and Bull are recorded values**, saved with each scenario active; switching `B62` recalculates them.
- **Forecast drivers are assumptions.** The tax rate is the fiscal 2021-2025 average excluding fiscal 2024's
  one-off EU State Aid charge; the sanity bands on forecasts are wide, so passing them is a weak test.
- **One company so far.** Safaricom and Equity Group, both listed in Nairobi, are the next models.
- **The fiscal 2026 first-quarter statements** are in `Source_filings/` but not yet used.

## Repository structure

```
Apple/Apple_Financial_Model.xlsx   the model: assumptions, statements, DCF, validation tab
Apple/Source_filings/              the 10-K and press-release PDFs behind the historicals
Apple/README.md                    model-specific assumptions and tabs
validate_model.py                  independent validator (statements, valuation, reverse DCF)
tests/                             fault-injection tests for the validator
images/                            screenshots of the model
docs/METHODOLOGY.md                full method and the review history
```

## How to run

```bash
pip install -r requirements.txt
python validate_model.py      # rebuild the statements and the valuation, report the implied WACC
python -m pytest              # prove the validator catches injected faults
```

Or open `Apple/Apple_Financial_Model.xlsx` in Excel or LibreOffice Calc and change the scenario in cell `B62` of
the Assumptions sheet.

## Documentation

The Apple model's tabs and assumptions are described in [`Apple/README.md`](Apple/README.md) and the full method in
[`docs/METHODOLOGY.md`](docs/METHODOLOGY.md).

## License

MIT. See [`LICENSE`](LICENSE). Sources: SEC EDGAR filings (CIK 0000320193), Apple investor relations, and
stockanalysis.com for the analyst range.

Alven Yuka · [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · [Email](mailto:alvenyuka2@gmail.com)
