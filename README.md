# Financial-Analyst

> Three-statement models, DCF valuations, and scenario analysis built from primary-source filings. Every historical line ties to the 10-K it came from.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#modelling-conventions)
[![Models](https://img.shields.io/badge/models-1%20%C2%B7%20Apple-blue)](#current-models)

![Financial-Analyst banner — primary-source three-statement models and DCF valuations](banner.svg)

## Why?

Most public financial models on the internet are detached from their source filings — figures appear without an audit trail, scenarios are toggled without traceability, and the balance sheet is left to drift. This repository publishes models where every historical line reconciles back to the 10-K, 10-Q, or press release it came from, and the workbook refuses to be treated as correct until a validation tab confirms structural integrity (balance-sheet identity, cash-flow ties, inter-statement consistency).

I'm Alven, a CPA finalist (Kenya) awaiting ICPAK membership. This repository is where I publish modelling work as I build out the technical side of accounting and finance.

## Features

- 📑 **Primary-source sourcing** — every workbook ships alongside the filings it was built from
- 🔍 **Validation tab** reconciles each historical line to its source and runs the balance-sheet identity, cash-flow ties, and inter-statement consistency checks
- 🎚️ **Scenario toggles** — bear / base / bull, but only on questions where the answer changes materially across them
- 🎨 **Colour-coded inputs**, locked calculated cells, single source of truth for assumptions
- 📂 **No macros, no external feeds, no add-ins** — opens in Excel or LibreOffice Calc

## Tech Stack

| Layer | Tools |
|---|---|
| Spreadsheet | Microsoft Excel (LibreOffice Calc compatible) |
| Modelling | Native formulas only — no VBA, no add-ins |
| Source filings | SEC EDGAR 10-K / 10-Q / press releases |

## Installation

```bash
git clone https://github.com/alvenyuka/Financial-Analyst.git
cd Financial-Analyst
# Open Apple_Financial_Model.xlsx in Excel or LibreOffice Calc
```

## Usage

1. Open the model file in Excel or LibreOffice Calc.
2. Review the assumptions tab (colour-coded inputs).
3. Toggle scenarios where wired up.
4. Confirm the **Validation** tab is fully green before trusting any output.
5. Cross-check any historical line against the source filing in `Source_filings/`.

## Modelling conventions

A few conventions run through every analysis in this repository.

### One source of truth for inputs

A dedicated tab holds assumptions and inputs. Every other tab recalculates from it. Inputs are colour-coded; calculated cells are locked.

### Every historical figure traces to its source

Each workbook ships alongside the 10-K, 10-Q, and press releases the figures were taken from. The validation tab reconciles each historical line back to the filing it came from.

### Structural integrity is enforced

The validation tab runs the balance-sheet identity (`assets = liabilities + equity`), the cash flow ties (net income → retained earnings, depreciation → PPE), and inter-statement consistency checks. The workbook does not treat a model as correct until those checks pass.

### Scenario toggles only where they help

Bear, base, and bull scenarios are wired up on questions where the answer changes materially across them. Where it does not, the model stays single-case to keep the audit trail clean.

### Excel-native, no add-ins

Files open in Excel or LibreOffice Calc. No macros, no external data feeds, no add-ins.

## Repository structure

```
Financial-Analyst/
├── README.md
├── banner.svg
├── Apple_Financial_Model.xlsx     # three-statement model, DCF, scenario toggles
└── Source_filings/                # 10-K, 10-Q, press releases backing the model
```

## Current models

| Company | Ticker | Status | Workbook |
|---|---|---|---|
| Apple Inc. | AAPL | ✅ Complete | [`Apple_Financial_Model.xlsx`](Apple_Financial_Model.xlsx) |

Additional company analyses will follow the same structure.

## Roadmap

- [x] Apple Inc. (AAPL) — three-statement model, DCF, validation harness
- [ ] Microsoft (MSFT)
- [ ] Safaricom (SCOM.NR) — Nairobi Securities Exchange
- [ ] Equity Group Holdings (EQTY.NR)

If you'd like a particular company modelled in this style, open an issue.

## Contributing

If you spot an error in any model, [open an issue](https://github.com/alvenyuka/Financial-Analyst/issues). Reproduction steps and the offending cell reference make a fix possible.

## License

MIT — see [`LICENSE`](LICENSE).

## Connect

Built by **Alven Yuka** — CPA Finalist (Kenya), awaiting ICPAK membership.

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alvenyuka) · 🐙 [GitHub](https://github.com/alvenyuka) · 🛠️ Fiverr: `alvenemmanuel`
