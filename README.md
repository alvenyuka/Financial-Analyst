# Financial-Analyst

> Company financial models built from primary-source SEC filings: three-statement models, DCF valuations, a validation tab citing the filing and page behind every historical line, and a Python validator that re-derives the statements and traces each line to its source tab.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#how-i-build-these)
[![Validation](https://img.shields.io/badge/validation-50%2F50%20lines%20reconciled-success)](#portfolio-snapshot)
[![models validated](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml/badge.svg)](https://github.com/alvenyuka/Financial-Analyst/actions/workflows/ci.yml)

![Financial-Analyst banner: primary-source three-statement models and DCF valuations](banner.svg)

## Why?

I'm Alven, a CPA Finalist (Kenya), awaiting ICPAK membership. This repo is where I publish the technical side of finance: modelling, valuation, and scenario analysis. Every model here ships with a validation tab that cross-references each historical line against the primary filing it came from, checks the balance sheet balances to zero in every year (historical and projected), and ties the cash flow statement's ending cash to the next period's balance sheet. If a figure can't be traced back to a filing, it doesn't go in the model.

## Project Structure

```
Financial-Analyst/
├── Apple/                         # Complete: three-statement model + DCF
│   ├── Apple_Financial_Model.xlsx
│   ├── Source_filings/            # 10-K, 10-Q, press releases behind the model
│   └── README.md                  # Model-specific tabs, validation results, DCF summary
├── validate_model.py               # Re-derives every identity outside the spreadsheet
├── tests/
│   ├── test_validate_model.py      # Breaks a copy of the model, checks each fault is caught
│   └── xlsx_surgery.py             # Edits one cached cell without dropping the rest
├── .github/workflows/ci.yml        # Runs that validator on every push
├── requirements.txt                # openpyxl, for the validator only
├── banner.svg
├── LICENSE
└── README.md                      # This file, repo-wide conventions
```

Each company gets its own folder with the workbook, the primary-source filings behind it, and a README describing that model's specific assumptions and results. [`Apple/`](./Apple/) is the only complete one right now, see [Roadmap](#roadmap) for what's next.

## Quick Start

1. Open [`Apple/`](./Apple/), the one complete model, and follow its own Quick Start.
2. In short: open `Apple_Financial_Model.xlsx` in Excel or LibreOffice Calc, start on the **Dashboard** tab, then confirm the **Validation** tab is fully green before trusting any output.
3. Or check it without opening Excel at all: `pip install -r requirements.txt && python validate_model.py` re-derives every total independently. See [Checking it yourself](#checking-it-yourself).
4. Cross-check any historical line against the source filing in `Apple/Source_filings/`.

## Features

- **Single source of truth for inputs.** One Assumptions tab per model; every other tab recalculates from it. Inputs are colour-coded, calculated cells are locked.
- **Every historical figure carries a primary-filing citation**, down to the page number, in the Validation tab's Source column. FY23 onward is archived in `Apple/Source_filings/`; the FY21 and FY22 figures cite the FY2022 and FY2023 10-Ks, which are linked to SEC EDGAR rather than archived here. The workbook's own tick marks check tab against tab, not against a PDF.
- **A validation tab reports ✓ or ✗ per line**, plus structural checks: balance-sheet identity, cash-flow ties, inter-statement consistency, forecast sanity bands.
- **Scenario / sensitivity toggles** where the question benefits from them (e.g. WACC × terminal growth, revenue growth × operating margin).
- **Excel-native.** No macros, no external data feeds, no add-ins. Opens in Excel or LibreOffice Calc.

## Tech Stack

| Layer | Tools |
|---|---|
| Spreadsheet | Microsoft Excel (LibreOffice Calc compatible) |
| Math | Native spreadsheet functions only, no macros, no add-ins |
| Sourcing | SEC EDGAR filings (10-K / 10-Q), company investor-relations press releases |

## Sheet Structure

Every model in this repo follows the same tab layout. [`Apple/Apple_Financial_Model.xlsx`](./Apple/) is the concrete example:

| Tab | Contents |
|---|---|
| Dashboard | Headline metrics |
| Cover | Model scope and version |
| Assumptions | Single source of truth for every driver: segment growth, margins, WACC, terminal growth, tax rate |
| IS / BS / CFS | Income statement, balance sheet, cash flow statement: historical years plus a multi-year projection |
| Ratios | Derived ratios |
| DCF | Unlevered FCF build, WACC discounting, Gordon Growth terminal value, sensitivity tables |
| Validation | Line-by-line filing cross-reference, structural integrity checks, forecast sanity bands |
| Pivots / Data Refs | Supporting pivot tables and reference data |

## Validation Harness

Each validation tab carries a filing-and-page citation per hardcoded historical figure, reconciles that figure against the statement tab it was transcribed to, checks the balance sheet balances to zero in every year (historical and projected), ties cash-flow ending cash to the next period's balance sheet, and confirms net income and D&A match between the CFS and IS. Forecast years are additionally checked against pre-defined plausibility bands (revenue growth, margins, tax rate, CapEx %, liquidity) so a projection can't silently drift outside a defensible range.

## Checking it yourself

The Validation tab says "ALL CHECKS PASS". That tick is computed by the same
spreadsheet whose correctness is in question, so it is worth exactly as much as
the formulas behind it. `validate_model.py` exists so you do not have to take it
on trust:

```bash
pip install -r requirements.txt
python validate_model.py                       # defaults to the Apple model
python validate_model.py path/to/Model.xlsx    # any model with a Validation tab
```

This runs in CI on every push, which is possible here and not in the modelling
repos: the workbook is committed, it is a few hundred kilobytes, and the checks
are arithmetic rather than training, so there is no dataset to fetch and nothing
to fit.

**The validator is itself tested**, which matters more than it might sound. A
script that has only ever been run against a correct workbook proves nothing: it
would report "all checks pass" just as confidently if its comparisons were
inverted or it were reading the wrong rows. So `tests/` copies the real model,
breaks one specific thing, and asserts that specific thing is caught. CI runs
those first and the real model second.

It adds the line items up in Python and compares its own arithmetic against the
totals the workbook reports. It never reads a cell whose value is a checkmark.
Nineteen identities are re-derived this way: the income statement from segment
revenue down to net income, both sides of the balance sheet ending in assets
minus liabilities minus equity, the cash flow statement's three sections and its
roll-forward, and the cross-statement tie from one year's ending cash to the next
year's opening cash. It exits non-zero if any of them break, so it can run in CI.

**What it used to miss, and now does not.** Those identities are computed from
the Validation tab, which is a second transcription of the statements rather than
the statements themselves. So none of them ever opened `IS`, `BS` or `CFS`. With
every numeric cell on the income statement and balance sheet replaced by the
number 1, the script still printed "19 of 19 ... matched" and exited 0, and the
CI badge would have stayed green. It now also checks each of the 49 Validation
line items against the `IS`, `BS` or `CFS` cells it was transcribed from, five
years each, taking the source-cell mapping from the Validation tab's own column G
formulas and doing the comparison in Python. Both families have to pass. A test
injects exactly that fault, changing a figure on the IS and leaving the
Validation tab alone, and asserts the run fails while the identities still pass.

**Still not checked, so a green badge is not read as more than it is:** the
`DCF`, `Ratios`, `Assumptions` and `Dashboard` tabs. Nothing in CI validates the
valuation, and the DCF tab has known defects, listed in
[`Apple/README.md`](./Apple/README.md).

**Current result on the Apple model: 19 of 19 identities match and 49 of 49 line
items trace to their source tab**, plus one advisory flag. The advisory is that the cash flow statement's ending cash exceeds
the balance sheet's cash line by roughly $0.8B to $1.3B in FY21 through FY23,
while FY24 and FY25 agree exactly. Apple's cash flow statement historically
reconciled to "cash, cash equivalents and restricted cash" where the balance
sheet line excludes restricted cash, so this is most likely the model faithfully
reproducing each filing's own presentation rather than an error. It is reported
rather than suppressed because "most likely" is not "checked", and it does not
affect the roll-forward tie this repo actually claims, which passes in every year.

## Portfolio Snapshot

**Apple Inc. (AAPL)**, the one complete model:

| Metric | Value |
|---|---|
| Validation-tab lines reconciled across tabs | 50 / 50 (100%), 5 years each |
| Citation basis for the historicals | SEC EDGAR (CIK 0000320193) + Apple investor-relations press releases, cited per line with a page number |
| Filings covered | 10-Ks filed Oct 2022, Nov 2023, Oct 2025 |
| Scenario the workbook is saved on | Bull (`Assumptions!B62`) |
| WACC and terminal growth the model computes | 7.79% and 3.0% (`Assumptions!B54`, `B55`) |
| Implied share price vs. the reference price | $240.05 vs. $232.50, about 3% upside |
| The same model at WACC 8.5% and g 2.5% | $205.92 (`DCF!D36`), about 11% downside |

Earlier versions of this table, and of the profile README, called WACC 8.5% with
terminal growth 2.5% the base case behind the $240. It is not: those are not the
inputs the model runs and not any scenario in its scenario engine. The
workbook's own caption says 8.5% / 2.5%, which is where the error came from. Both
figures are shown above because the direction of the call changes between them.

The DCF tab has other known defects, including a terminal value discounted one
period too many and a football-field row that reads from a different model. They
are listed in full in [`Apple/README.md`](./Apple/README.md), and none of them is
checked by the validator, which does not open the DCF tab.

Full detail, including the two DCF sensitivity tables (WACC × terminal growth, revenue growth × operating margin), is in [`Apple/README.md`](./Apple/README.md#validation).

## Roadmap

- [x] Repo conventions: single-source assumptions tab, filing-traced historicals, validation tab, sensitivity toggles
- [x] Apple Inc. (AAPL): three-statement model + DCF, 50/50 lines reconciled
- [ ] Correct the DCF tab's known defects and restate the valuation (see `Apple/README.md`)
- [ ] Extend the validator to the DCF, Ratios and Assumptions tabs
- [ ] Microsoft (MSFT)
- [ ] Safaricom (SCOM.NR), Nairobi Securities Exchange
- [ ] Equity Group Holdings (EQTY.NR)

## License

MIT. See [`LICENSE`](LICENSE).

## Credits

Author: **Alven Yuka**, CPA Finalist (Kenya). Built from primary-source SEC EDGAR filings and company investor-relations disclosures; see each model's own Credits/Validation section for its specific sources.

## Connect

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)

## Contributing

If you spot an error in any model, [open an issue](https://github.com/alvenyuka/Financial-Analyst/issues) with the offending cell reference.
