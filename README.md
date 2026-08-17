# Financial-Analyst

> Company financial models built from primary-source SEC filings: three-statement models, DCF valuations, and a validation tab that ties every historical line back to the 10-K it came from.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#how-i-build-these)
[![Validation](https://img.shields.io/badge/validation-50%2F50%20data%20points-success)](#portfolio-snapshot)

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
├── banner.svg
├── LICENSE
└── README.md                      # This file, repo-wide conventions
```

Each company gets its own folder with the workbook, the primary-source filings behind it, and a README describing that model's specific assumptions and results. [`Apple/`](./Apple/) is the only complete one right now, see [Roadmap](#roadmap) for what's next.

## Quick Start

1. Open [`Apple/`](./Apple/), the one complete model, and follow its own Quick Start.
2. In short: open `Apple_Financial_Model.xlsx` in Excel or LibreOffice Calc, start on the **Dashboard** tab, then confirm the **Validation** tab is fully green before trusting any output.
3. Cross-check any historical line against the source filing in `Apple/Source_filings/`.

## Features

- **Single source of truth for inputs.** One Assumptions tab per model; every other tab recalculates from it. Inputs are colour-coded, calculated cells are locked.
- **Every historical figure traces to a primary filing** archived in the same folder as the model, down to the page number.
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

Each validation tab cross-references every hardcoded historical figure against the company's own SEC filings and press releases, checks the balance sheet balances to zero in every year (historical and projected), ties cash-flow ending cash to the next period's balance sheet, and confirms net income and D&A match between the CFS and IS. Forecast years are additionally checked against pre-defined plausibility bands (revenue growth, margins, tax rate, CapEx %, liquidity) so a projection can't silently drift outside a defensible range.

## Portfolio Snapshot

**Apple Inc. (AAPL)**, the one complete model:

| Metric | Value |
|---|---|
| Historical data points verified against 10-K | 50 / 50 (100%) |
| Verification basis | SEC EDGAR (CIK 0000320193) + Apple investor-relations press releases |
| Filings covered | 10-Ks filed Oct 2022, Nov 2023, Oct 2025 |
| DCF base case | WACC 8.5%, terminal growth 2.5% |
| Implied share price vs. reference price at build time | ≈$240 vs. $232.50, roughly 3% upside |

Full detail, including the two DCF sensitivity tables (WACC × terminal growth, revenue growth × operating margin), is in [`Apple/README.md`](./Apple/README.md#validation).

## Roadmap

- [x] Repo conventions: single-source assumptions tab, filing-traced historicals, validation tab, sensitivity toggles
- [x] Apple Inc. (AAPL): three-statement model + DCF, 50/50 data points verified
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
