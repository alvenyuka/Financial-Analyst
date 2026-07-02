# Financial-Analyst

> Company financial models built from primary-source SEC filings — three-statement models, DCF valuations, and a validation tab that ties every historical line back to the 10-K it came from.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-LibreOffice%20compatible-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![No macros](https://img.shields.io/badge/macros-none-success)](#how-i-build-these)

![Financial-Analyst banner — primary-source three-statement models and DCF valuations](banner.svg)

## Why?

Most public financial models are detached from their source filings — figures appear without an audit trail, scenarios are toggled without traceability. Every model in this repo ships with a validation tab that cross-references each historical line against the primary filing, checks the balance sheet balances to zero in every year (historical and projected), and ties the cash flow statement's ending cash to the next period's balance sheet.

I'm Alven, a CPA Finalist (Kenya), awaiting ICPAK membership. This repo is where I publish the technical side of finance: modelling, valuation, scenario analysis, and the data discipline that makes any of it credible.

## What's inside

| Company | Status | Folder |
|---|---|---|
| Apple Inc. (AAPL) | Complete — 50/50 historical data points verified against 10-K, all structural checks pass | [`Apple/`](./Apple/) |
| Microsoft (MSFT) | Planned | — |
| Safaricom (SCOM.NR) — Nairobi Securities Exchange | Planned | — |
| Equity Group Holdings (EQTY.NR) | Planned | — |

Each project folder contains the workbook, the primary-source filings behind it, and a README describing the model's specific assumptions and results.

## How I build these

Same conventions across every analysis here:

- One tab for inputs and assumptions. Every other tab recalculates from it — colour-coded inputs, locked calculated cells.
- Every historical figure traces to a primary filing archived in the same folder as the model.
- A validation tab ties each historical line back to the source and reports ✓ or ✗, plus structural checks: balance-sheet identity, cash-flow ties, inter-statement consistency.
- Scenario / sensitivity toggles where the question benefits from them (e.g. WACC × terminal growth, revenue growth × operating margin).

Files open in Excel or LibreOffice Calc. No macros, no external data feeds, no add-ins.

## Contributing

If you spot an error in any model, [open an issue](https://github.com/alvenyuka/Financial-Analyst/issues) with the offending cell reference.

## License

MIT — see [`LICENSE`](LICENSE).

## Connect

[alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · [GitHub](https://github.com/alvenyuka) · Fiverr: alvenemmanuel

If you'd like work in this style for your own company, [open an issue](https://github.com/alvenyuka/Financial-Analyst/issues) or send a note.
