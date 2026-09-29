# Financial-Analyst

> Can every number in a financial model be traced back to the filing it came from? Three-statement models and DCF valuations built from SEC filings, with an independent Python check that re-derives the statements on every change: 19 of 19 accounting identities and 49 of 49 line items reconcile on the Apple model.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#how-it-works)
[![models validated](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml/badge.svg)](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml)

![Financial-Analyst banner: primary-source three-statement models and DCF valuations](banner.svg)

## The problem

Investment, credit and audit decisions often rest on spreadsheet models, and a single mistyped historical
figure flows silently into every forecast and valuation built on it. A model's own "all checks pass" cell
is computed by the same formulas it is supposed to check, so it proves little.

These models are built the way an auditor would want to review them: every historical figure cites the
filing and page it came from, and a separate program re-checks the arithmetic outside the spreadsheet.

## What I found

On the Apple Inc. model (fiscal years 2021 to 2025):

| Check | Result |
|---|---:|
| Accounting identities re-derived outside Excel (income statement, balance sheet, cash flow, cross-statement ties) | **19 of 19** |
| Line items traced back to the statement they were transcribed from, five years each | **49 of 49** |
| DCF implied share price vs reference price, at the model's WACC of 7.79% and terminal growth of 3.0% | $240.05 vs $232.50 |
| The same model at WACC 8.5% and terminal growth 2.5% | $205.92 |

- **The accounting holds.** Every total the workbook reports agrees with an independent recalculation, and
  every validation line matches its source statement.
- **The valuation call depends on two assumptions.** Moving WACC from 7.79% to 8.5% and terminal growth from
  3.0% to 2.5% turns about 3% upside into about 11% downside. A valuation this sensitive should be presented
  as a range, not a single price.
- **One presentation difference is flagged, not hidden.** In FY21 to FY23 cash-flow ending cash exceeds the
  balance-sheet cash line by about $0.8B to $1.3B, most likely because the cash flow includes restricted
  cash. It is reported for review rather than suppressed.

**What I would tell an investment committee:** rely on the historical statements, which are fully traced and
reconciled; treat the DCF as a sensitivity range rather than a target price until its known defects are
fixed.

## How it works

1. **One assumptions tab per model** drives every other tab; inputs are colour-coded and calculated cells
   locked.
2. **Historical figures from primary sources**: SEC 10-K and 10-Q filings and Apple investor-relations
   releases, each cited to a page in the Validation tab.
3. **Three statements and a DCF**: unlevered free cash flow, WACC discounting, Gordon Growth terminal value,
   and sensitivity tables.
4. **An independent validator** (`validate_model.py`) adds the line items up in Python and compares them with
   the workbook's totals. Its own tests break a copy of the model on purpose and confirm each fault is caught.
   It runs on every push.

## Run it

```bash
pip install -r requirements.txt
python validate_model.py          # checks the Apple model; pass a path to check another
python -m pytest                  # confirms the validator catches deliberately broken models
```

Or open `Apple/Apple_Financial_Model.xlsx` in Excel or LibreOffice Calc and start on the Dashboard tab.

## Limitations

- **The DCF tab has known defects**, including a terminal value discounted one period too many, so the
  implied price is not yet a reliable call. Fixing and restating it is the next step.
- **The validator does not check the valuation**: the DCF, Ratios, Assumptions and Dashboard tabs are outside
  its scope.
- **One complete model so far** (Apple). Safaricom and Equity Group, both listed in Nairobi, are planned.

## More detail

Model-specific assumptions, the DCF defects and the sensitivity tables are in
[`Apple/README.md`](Apple/README.md). The full write-up of the conventions and the validator is in
[`docs/METHODOLOGY.md`](docs/METHODOLOGY.md).

## License

MIT. See [`LICENSE`](LICENSE). Sources: SEC EDGAR filings (CIK 0000320193) and Apple investor-relations disclosures.

## Connect

Built by Alven Yuka, CPA Finalist and Accounting Specialist at GIZ, Nairobi.

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)
