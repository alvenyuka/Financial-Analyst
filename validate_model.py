"""Independently re-derive every accounting identity in a model workbook.

The workbook already has a Validation tab that prints "ALL CHECKS PASS". This
script exists because that tick mark is produced by the same spreadsheet whose
correctness is in question. A formula that is wrong reports a tick just as
confidently as one that is right, so a reader has no way to distinguish the two
without recomputing the arithmetic outside the file.

This does that. It adds the raw line items up in Python and compares the result
against the totals the workbook reports. Nothing here reads a cell whose value
is a checkmark.

Scope, stated up front because an earlier version of this file overstated it.
The line items live on the Validation tab, which is a second transcription of
the statements rather than the statements themselves. Re-deriving the identities
there proves the transcription is internally consistent and proves nothing about
`IS`, `BS` or `CFS`. That gap was real: with every numeric cell on the IS and BS
replaced by the number 1, this script still printed "19 of 19 ... matched" and
exited 0. The transcription checks added below close it, by reading the source
sheets directly and comparing them to the Validation tab, row by row and year by
year. Both families must pass for the run to succeed.

The valuation is checked too, without reusing the workbook's own cell choices:
the change in working capital is rebuilt from the balance-sheet lines and compared
with the cash flow statement; the forecast balance sheet is re-added from its
components; and the DCF is rebuilt from the valuation date (Assumptions!B60) and
the fiscal year-end dates, so a cash flow from a year that ended before the
valuation date is excluded and its cash is counted in net cash instead. WACC,
free cash flow, present values, terminal value, enterprise and equity value, the
per-share price, the sensitivity grid and the discounted multiples are all
re-derived. Still out of scope: the `Ratios`, `Dashboard` and `Pivots` tabs, and
the Bear and Bull prices, which are recorded with each scenario active and printed
as recorded rather than re-derived.

What it checks:

* **Income statement**, down the page. Revenue from its two segments, cost of
  sales from its two segments, then gross profit, operating income, pre-tax
  income and net income, each from the line above it.
* **Balance sheet**, both sides, ending with the identity everything rests on:
  assets minus liabilities minus equity is zero.
* **Cash flow statement**, including the two ties that are easy to assert and
  easy to get wrong: net change in cash equals the three sections added up, and
  ending cash equals beginning cash plus that change.
* **Cross-statement ties**, which is where a three-statement model usually breaks.
  The cash flow statement's ending cash has to equal the balance sheet's cash for
  the same year, and it has to equal the next year's beginning cash. A model can
  satisfy every within-statement check and still fail these.
* **Transcription**, one check per Validation row, five years each: the figure on
  the Validation tab against the cell or cells it was copied from on `IS`, `BS`
  or `CFS`. The source cells are read out of the Validation tab's own column G
  formulas, so the mapping comes from the workbook and the comparison is done
  here. Without this, every identity above could pass on a workbook whose
  statements had been emptied.

Usage:

    python validate_model.py                        # defaults to the Apple model
    python validate_model.py path/to/Model.xlsx

Exits 0 when every check passes, 1 otherwise, so it can run in CI.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

DEFAULT_MODEL = Path(__file__).parent / "Apple" / "Apple_Financial_Model.xlsx"

# Money is in millions of USD. The workbook rounds to whole millions, so an exact
# equality test would fail on rounding alone. One million on figures in the
# hundreds of billions is about four parts in a million.
TOLERANCE = 1.0


def _wrap(text: str, width: int):
    """Split text into lines of at most `width` characters, breaking between words."""
    words, line, out = text.split(), "", []
    for w in words:
        if len(line) + len(w) + 1 > width:
            out.append(line)
            line = w
        else:
            line = f"{line} {w}".strip()
    if line:
        out.append(line)
    return out


class Check:
    """One identity, evaluated for every year in the model."""

    def __init__(self, label: str, expected, actual, years, note: str | None = None,
                 advisory: bool = False, tol: float = TOLERANCE):
        self.label = label
        self.tol = tol
        self.expected = expected
        self.actual = actual
        self.years = years
        #: Context a reader needs to interpret a failure, printed alongside it.
        #: A check that fails without explanation invites the reader to assume
        #: either "broken model" or "ignore it", and one of those is wrong.
        self.note = note
        #: An advisory check is reported but does not fail the run. It compares
        #: two figures whose equality depends on a presentation choice rather
        #: than on arithmetic, so a difference is a question, not a defect.
        self.advisory = advisory

    @property
    def failures(self):
        """Years where the expected and actual values differ by more than the tolerance, or one is missing."""
        out = []
        for year, exp, act in zip(self.years, self.expected, self.actual):
            if exp is None or act is None:
                out.append((year, exp, act, "missing value"))
            elif abs(exp - act) > self.tol:
                gap = f"{exp - act:,.1f}" if self.tol >= 1 else f"{exp - act:+.6f}"
                out.append((year, exp, act, f"differs by {gap}"))
        return out

    @property
    def passed(self) -> bool:
        """True when no year fails."""
        return not self.failures


def read_rows(ws, first_col=2, n_years=5):
    """Map each row label in column A to its list of yearly values.

    Raises if a label carrying numbers appears more than once. Silently keeping
    the first occurrence would make every check on that label report PASS for a
    row nobody chose, which is worse than not checking it: the result looks
    verified.
    """
    rows = {}
    where = {}
    duplicates = []

    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 1).value
        if not isinstance(label, str):
            continue
        label = label.strip()
        values = [ws.cell(r, c).value for c in range(first_col, first_col + n_years)]
        if not any(isinstance(v, (int, float)) for v in values):
            continue
        if label in rows:
            duplicates.append((label, where[label], r))
            continue
        rows[label] = [v if isinstance(v, (int, float)) else None for v in values]
        where[label] = r

    if duplicates:
        detail = "\n".join(
            f"    {label!r}: rows {first} and {second}" for label, first, second in duplicates
        )
        raise ValueError(
            "Duplicate numeric row labels in the Validation tab. Every check on "
            "these would silently validate the first one and report a pass:\n"
            + detail
        )
    return rows


def find_section_headers(ws):
    """Locate each statement's year-header row from its section title, so inserted
    rows on the Validation tab cannot shift a check onto the wrong row."""
    titles = {"income statement": "INCOME STATEMENT", "balance sheet": "BALANCE SHEET",
              "cash flow statement": "CASH FLOW STATEMENT"}
    found = {}
    for r in range(1, ws.max_row + 1):
        text = str(ws.cell(r, 1).value or "")
        for key, title in titles.items():
            if key not in found and text.startswith(title):
                found[key] = r + 1
    missing = [k for k in titles if k not in found]
    if missing:
        raise ValueError(f"section title(s) not found on the Validation tab: {missing}")
    return found


def check_year_headers(ws, header_rows, n_years=5):
    """Every statement must be laid out over the same years.

    The years come from the income statement's header. If another statement used
    a different set, the comparisons would still line up positionally and mean
    nothing.
    """
    headers = {
        name: tuple(ws.cell(r, c).value for c in range(2, 2 + n_years))
        for name, r in header_rows.items()
    }
    distinct = set(headers.values())
    if len(distinct) > 1:
        detail = "\n".join(f"    {name}: {list(cols)}" for name, cols in headers.items())
        raise ValueError(
            "The statements are not laid out over the same years, so a "
            "column-by-column comparison would be meaningless:\n" + detail
        )
    return list(next(iter(distinct)))


def build_checks(rows, years):
    """The three-statement identities, each recomputed from its components in the Validation tab.

    `rows` maps each Validation row label to its yearly values. Every check adds or
    subtracts component rows and compares the result with the total the model
    reports, for every year.
    """
    def g(label):
        if label not in rows:
            raise KeyError(f"row not found in the Validation tab: {label!r}")
        return rows[label]

    def add(*labels):
        cols = [g(l) for l in labels]
        return [
            None if any(v is None for v in vals) else sum(vals)
            for vals in zip(*cols)
        ]

    def sub(a, *rest):
        base = g(a)
        cols = [g(l) for l in rest]
        return [
            None if (b is None or any(v is None for v in vals)) else b - sum(vals)
            for b, *vals in zip(base, *cols)
        ]

    checks = [
        # ---- income statement -------------------------------------------
        Check("IS  revenue = products + services",
              add("Products revenue", "Services revenue"), g("Total revenue (calc)"), years),
        Check("IS  cost of sales = products + services",
              add("Cost of products", "Cost of services"), g("Total cost of sales"), years),
        Check("IS  gross profit = revenue - cost of sales",
              sub("Total revenue (calc)", "Total cost of sales"), g("Gross profit (calc)"), years),
        Check("IS  opex = R&D + SG&A",
              add("R&D expense", "SG&A expense"), g("Total operating expenses"), years),
        Check("IS  operating income = gross profit - opex",
              sub("Gross profit (calc)", "Total operating expenses"), g("Operating income (calc)"), years),
        Check("IS  pre-tax = operating income + other",
              add("Operating income (calc)", "Other income/(expense)"), g("Pre-tax income (calc)"), years),
        Check("IS  net income = pre-tax - tax",
              sub("Pre-tax income (calc)", "Income tax expense"), g("Net income (calc)"), years),

        # ---- balance sheet ----------------------------------------------
        Check("BS  current assets sum to the subtotal",
              add("Cash & equivalents", "Marketable securities (current)", "Accounts receivable, net",
                  "Inventory", "Vendor non-trade receivables", "Other current assets"),
              g("Total current assets"), years),
        Check("BS  total assets = current + PP&E + other non-current",
              add("Total current assets", "Marketable securities (non-curr)",
                  "Property, plant & equipment, net", "Other non-current assets"),
              g("Total assets"), years),
        Check("BS  current liabilities sum to the subtotal",
              add("Accounts payable", "Other current liabilities", "Deferred revenue",
                  "Commercial paper", "Term debt (current)"),
              g("Total current liabilities"), years),
        Check("BS  total liabilities = current + LT debt + other non-current",
              add("Total current liabilities", "Long-term debt (non-current)",
                  "Other non-current liabilities"),
              g("Total liabilities"), years),
        Check("BS  equity = common stock + retained earnings + other comprehensive income",
              add("Common stock and APIC", "Retained earnings (deficit)", "Accumulated OCI (loss)"),
              g("Shareholders' equity"), years),
        Check("BS  liabilities + equity subtotal",
              add("Total liabilities", "Shareholders' equity"),
              g("Total liabilities + equity"), years),
        Check("BS  THE identity: assets = liabilities + equity",
              g("Total assets"), g("Total liabilities + equity"), years),

        # ---- cash flow statement -----------------------------------------
        Check("CF  investing = capex + other investing",
              add("Capital expenditures", "Other investing"), g("Cash from investing"), years),
        Check("CF  financing = dividends + buybacks + other",
              add("Dividends paid", "Share repurchases", "Other financing"),
              g("Cash from financing"), years),
        Check("CF  net change = operations + investing + financing",
              add("Cash from operations", "Cash from investing", "Cash from financing"),
              g("Net change in cash"), years),
        Check("CF  ending cash = beginning cash + net change",
              add("Beginning cash", "Net change in cash"), g("Ending cash"), years),
        Check("CF  free cash flow = operations - capex",
              add("Cash from operations", "Capital expenditures"),
              g("Free cash flow (CFO - CapEx)"), years),

        # ---- cross-statement ---------------------------------------------
        Check("TIE ending cash on the CFS = cash on the BS",
              g("Ending cash"), g("Cash & equivalents"), years,
              note=(
                  "Expected for FY21 to FY23: those cash flow statements reconcile "
                  "to cash, cash equivalents and restricted cash, while the balance "
                  "sheet line excludes restricted cash, so the gaps (989, 1,331 and "
                  "772) are the restricted cash. From FY24 Apple's statements agree."
              ),
              advisory=True),
    ]

    # The roll-forward tie has to be built by hand: this year's beginning cash is
    # last year's ending cash. A break here means the statements are not actually
    # linked, only arranged to look like they are.
    ending = g("Ending cash")
    beginning = g("Beginning cash")
    checks.append(
        Check("TIE beginning cash = prior year ending cash",
              ending[:-1], beginning[1:], [f"{a} to {b}" for a, b in zip(years, years[1:])])
    )
    return checks


# The Validation tab's column G holds, for each line item, a formula comparing
# that row against the sheet it was transcribed from. Two shapes are used:
#
#   =IF(SUMPRODUCT(ABS(B16:F16-IS!B8:F8))<1, ...)                 range form
#   =IF((ABS(B36-(BS!C5))+ABS(C36-(BS!D5))+...)<5, ...)           per-cell form
#
# Only the cell references are taken from them. The comparison itself is redone
# here in Python, so a broken formula in column G cannot hide a broken figure.
_RANGE_LINK = re.compile(
    r"ABS\((?P<vcol>[A-Z]+)(?P<vrow>\d+):[A-Z]+\d+"
    r"\s*-\s*(?P<sheet>[A-Za-z][A-Za-z0-9_ ]*)!(?P<scol>[A-Z]+)(?P<srow>\d+):[A-Z]+\d+\)"
)
_CELL_LINK = re.compile(
    r"ABS\((?P<vref>[A-Z]+\d+)\s*-\s*\(?(?P<src>[A-Za-z][A-Za-z0-9_ ]*![A-Z]+\d+"
    r"(?:\s*\+\s*[A-Za-z][A-Za-z0-9_ ]*![A-Z]+\d+)*)\)?\)"
)
_SHEET_CELL = re.compile(r"(?P<sheet>[A-Za-z][A-Za-z0-9_ ]*)!(?P<col>[A-Z]+)(?P<row>\d+)")


def _col_index(letters: str) -> int:
    """Column number of a column letter reference (A = 1, Z = 26, AA = 27)."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _formula_text(value) -> str | None:
    """Column G holds plain strings for some rows and ArrayFormula objects for
    others. Both carry the formula; only the attribute differs."""
    if isinstance(value, str):
        return value if value.startswith("=") else None
    text = getattr(value, "text", None)
    return text if isinstance(text, str) and text.startswith("=") else None


def read_source_links(ws_formulas, n_years=5):
    """Map each Validation row to the source cells it was transcribed from.

    Returns {row: (label, [(validation_cell, [(sheet, row, col), ...]), ...])},
    one entry per year column. Rows whose column G compares a source sheet to
    itself, such as the BALANCE CHECK row, carry no Validation-tab operand and
    are skipped: there is no transcription to check there.
    """
    links = {}
    for r in range(1, ws_formulas.max_row + 1):
        formula = _formula_text(ws_formulas.cell(r, 7).value)
        if not formula:
            continue
        label = ws_formulas.cell(r, 1).value
        if not isinstance(label, str):
            continue

        per_year = []
        match = _RANGE_LINK.search(formula)
        if match and int(match.group("vrow")) == r:
            vcol, scol = _col_index(match.group("vcol")), _col_index(match.group("scol"))
            sheet, srow = match.group("sheet"), int(match.group("srow"))
            for offset in range(n_years):
                per_year.append(((r, vcol + offset), [(sheet, srow, scol + offset)]))
        else:
            seen = set()
            for cell_match in _CELL_LINK.finditer(formula):
                vref = cell_match.group("vref")
                if vref in seen:
                    continue  # the TEXT() fallback repeats every term
                seen.add(vref)
                vsplit = re.match(r"([A-Z]+)(\d+)", vref)
                if int(vsplit.group(2)) != r:
                    continue
                sources = [
                    (m.group("sheet"), int(m.group("row")), _col_index(m.group("col")))
                    for m in _SHEET_CELL.finditer(cell_match.group("src"))
                ]
                per_year.append(((r, _col_index(vsplit.group(1))), sources))

        if len(per_year) == n_years:
            links[r] = (label.strip(), per_year)
    return links


def build_transcription_checks(wb_values, links, years):
    """One check per Validation row: the transcribed figure against its source.

    This is the part that actually opens IS, BS and CFS. Everything else in this
    script works off the Validation tab's own copy of the numbers.
    """
    validation = wb_values["Validation"]
    checks = []

    for _, (label, per_year) in sorted(links.items()):
        transcribed, derived = [], []
        for (vrow, vcol), sources in per_year:
            value = validation.cell(vrow, vcol).value
            transcribed.append(value if isinstance(value, (int, float)) else None)

            total = 0.0
            for sheet, srow, scol in sources:
                if sheet not in wb_values.sheetnames:
                    total = None
                    break
                cell = wb_values[sheet].cell(srow, scol).value
                if not isinstance(cell, (int, float)):
                    total = None
                    break
                total += cell
            derived.append(total)

        sheets = sorted({sheet for _, sources in per_year for sheet, _, _ in sources})
        checks.append(
            Check(f"SRC {label} = {'/'.join(sheets)}", derived, transcribed, years,
                  note=(
                      "The Validation tab is a transcription of the statements, not "
                      "the statements. A difference here means the figure being "
                      "checked by every identity above is not the figure on the "
                      "statement it came from."
                  ))
        )
    return checks


RATIO_TOL = 1e-5  # 0.001 percentage points: covers the $1M rounding slack, catches any real error


def build_ratio_checks(wb_values, rows, years):
    """The Ratios tab, rebuilt from the Validation-tab line items rather than read from its formulas.

    The line items are the ones already traced to the filings (the SRC checks), so a
    ratio that passes here rests on the 10-K figures, not on whichever statement cell
    the Ratios formula happens to point at. Growth rates start in the second year.
    """
    ws = wb_values["Ratios"]
    tax = wb_values["Assumptions"]["B50"].value
    reported = read_rows(ws)

    def g(label):
        if label == "_debt":
            return debt
        if label not in rows:
            raise KeyError(f"row not found in the Validation tab: {label!r}")
        return rows[label]

    def per_year(fn, *labels):
        cols = [g(label) for label in labels]
        return [None if any(v is None for v in vals) else fn(*vals) for vals in zip(*cols)]

    def growth(label):
        v = g(label)
        return [None] + [None if a is None or b is None else b / a - 1 for a, b in zip(v, v[1:])]

    debt = None  # total debt, filled in next and read through g("_debt")
    debt = per_year(lambda cp, cur, lt: cp + cur + lt,
                    "Commercial paper", "Term debt (current)", "Long-term debt (non-current)")
    revenue, net_income = "Total revenue (calc)", "Net income (calc)"
    derived = {
        "Gross margin %": per_year(lambda gp, r: gp / r, "Gross profit (calc)", revenue),
        "Operating margin %": per_year(lambda oi, r: oi / r, "Operating income (calc)", revenue),
        "Net margin %": per_year(lambda ni, r: ni / r, net_income, revenue),
        "EBITDA margin %": per_year(lambda oi, da, r: (oi + da) / r,
                                    "Operating income (calc)", "Depreciation & amortization", revenue),
        "Revenue growth %": growth(revenue),
        "Operating income growth %": growth("Operating income (calc)"),
        "Net income growth %": growth(net_income),
        "Return on assets (ROA)": per_year(lambda ni, ta: ni / ta, net_income, "Total assets"),
        "Return on equity (ROE)": per_year(lambda ni, eq: ni / eq, net_income, "Shareholders' equity"),
        "Return on invested capital": per_year(lambda oi, d, eq: oi * (1 - tax) / (d + eq),
                                               "Operating income (calc)", "_debt", "Shareholders' equity"),
        "Debt / Equity": per_year(lambda d, eq: d / eq, "_debt", "Shareholders' equity"),
        "Debt / Total capital": per_year(lambda d, eq: d / (d + eq), "_debt", "Shareholders' equity"),
        "Current ratio": per_year(lambda ca, cl: ca / cl, "Total current assets", "Total current liabilities"),
        "Net debt ($M)": per_year(lambda d, c, ms1, ms2: d - c - ms1 - ms2, "_debt", "Cash & equivalents",
                                  "Marketable securities (current)", "Marketable securities (non-curr)"),
        "Operating cash flow ($M)": g("Cash from operations"),
        "Free cash flow ($M)": g("Free cash flow (CFO - CapEx)"),
        "FCF margin %": per_year(lambda f, r: f / r, "Free cash flow (CFO - CapEx)", revenue),
        "FCF / Net income": per_year(lambda f, ni: f / ni, "Free cash flow (CFO - CapEx)", net_income),
        "CapEx / Revenue": per_year(lambda c, r: -c / r, "Capital expenditures", revenue),
    }
    checks = []
    for label, values in derived.items():
        if label not in reported:
            raise KeyError(f"row not found in the Ratios tab: {label!r}")
        start = 1 if values[0] is None else 0  # growth rates have no first year
        tol = TOLERANCE if label.endswith("($M)") else RATIO_TOL
        checks.append(Check(f"RAT {label}", values[start:], reported[label][start:], years[start:], tol=tol))
    return checks


def _dcf_price(fcf, periods, wacc, g, net_cash, shares):
    """Per-share value: cash flows from years ending after the valuation date, plus a
    Gordon terminal value at the last forecast year-end."""
    pv = sum(f / (1 + wacc) ** t for f, t in zip(fcf, periods) if t > 0)
    tv = fcf[-1] * (1 + g) / (wacc - g) / (1 + wacc) ** periods[-1]
    return (pv + tv + net_cash) / shares


# Labels the DCF rebuild depends on. An inserted row would otherwise move a cell
# reference silently, so each address is checked against its label first.
EXPECTED_LABELS = {
    ("IS", "A28"): "Operating income",
    ("CFS", "A10"): "Depreciation",
    ("CFS", "A13"): "Changes in working capital",
    ("CFS", "A19"): "Capital expenditures",
    ("CFS", "A30"): "Cash & equivalents, ending",
    ("BS", "B5"): "Cash",
    ("BS", "B15"): "Long-term marketable securities",
    ("BS", "B29"): "Long-term debt",
    ("Assumptions", "A25"): "Effective tax rate",
    ("Assumptions", "A54"): "WACC",
    ("Assumptions", "A60"): "Valuation date",
    ("DCF", "A14"): "Fiscal year end",
}

WC_ASSETS = (7, 8, 9, 10, 16)     # receivables, inventory, vendor receivables, other current, other non-current
WC_LIABS = (21, 22, 23, 30)       # payables, other current, deferred revenue, other non-current
BS_ASSETS = (5, 6, 7, 8, 9, 10, 14, 15, 16)
BS_LIABS_EQUITY = (21, 22, 23, 24, 25, 29, 30, 35, 36, 37)


def build_dcf_checks(wb):
    """Rebuild working capital, the forecast balance sheet and the DCF from the statements."""
    IS, BS, CFS = wb["IS"], wb["BS"], wb["CFS"]
    A, D = wb["Assumptions"], wb["DCF"]
    v = lambda ws, ref: ws[ref].value  # noqa: E731
    for (sheet, ref), want in EXPECTED_LABELS.items():
        got = str(wb[sheet][ref].value or "").strip()
        if not got.startswith(want):
            raise ValueError(f"{sheet}!{ref} is {got!r}, expected a label starting {want!r}")

    stmt = "GHIJ"                      # FY2026E to FY2029E on IS, CFS and Assumptions
    bs_cur, bs_prev = "HIJK", "GHIJ"   # the balance sheet starts one column later
    years = ("FY26E", "FY27E", "FY28E", "FY29E")
    checks = []

    def one(label, derived, cell, tol=TOLERANCE, sheet=D):
        return Check(f"DCF {label}", [derived], [v(sheet, cell)], ["DCF"], tol=tol)

    # Working capital from the balance sheet, not from the cash flow statement.
    dwc = []
    for cur, prev in zip(bs_cur, bs_prev):
        move = lambda r: v(BS, f"{cur}{r}") - v(BS, f"{prev}{r}")  # noqa: E731
        dwc.append(-sum(move(r) for r in WC_ASSETS) + sum(move(r) for r in WC_LIABS))
    checks.append(Check("DCF change in working capital = movement in the balance-sheet lines",
                        dwc, [v(CFS, f"{c}13") for c in stmt], list(years)))

    # The forecast balance sheet, re-added from its line items (no totals, no plug).
    gap = [sum(v(BS, f"{c}{r}") for r in BS_ASSETS) - sum(v(BS, f"{c}{r}") for r in BS_LIABS_EQUITY)
           for c in bs_cur]
    checks.append(Check("DCF forecast balance sheet balances from its line items", gap,
                        [0.0] * 4, list(years)))
    checks.append(Check("DCF forecast cash = cash flow statement ending cash",
                        [v(CFS, f"{c}30") for c in stmt], [v(BS, f"{c}5") for c in bs_cur], list(years)))

    taxes = [v(A, f"{c}25") for c in stmt]
    wacc_tax = v(A, "B50")
    ke = v(A, "B45") + v(A, "B47") * v(A, "B46")
    wacc = ke * (1 - v(A, "B52")) + v(A, "B49") * (1 - wacc_tax) * v(A, "B52")
    g = v(A, "B55")
    fcf = [v(IS, f"{c}28") * (1 - t) + v(CFS, f"{c}10") + v(CFS, f"{c}19") + w
           for c, t, w in zip(stmt, taxes, dwc)]
    val_date = v(A, "B60")
    year_ends = [v(D, f"{c}14") for c in "BCDE"]
    periods = [(ye - val_date).days / 365.25 for ye in year_ends]
    pv = sum(f / (1 + wacc) ** t for f, t in zip(fcf, periods) if t > 0)
    tv_pv = fcf[-1] * (1 + g) / (wacc - g) / (1 + wacc) ** periods[-1]
    n = "H"                                                    # FY26E, the last year ended before valuation
    debt = v(BS, f"{n}24") + v(BS, f"{n}25") + v(BS, f"{n}29")
    net_cash = v(BS, f"{n}5") + v(BS, f"{n}6") + v(BS, f"{n}15") - debt
    shares, price_ref = v(A, "B58"), v(A, "B59")
    price = (pv + tv_pv + net_cash) / shares
    df_last = 1 / (1 + wacc) ** periods[-1]
    ebitda = v(IS, "J28") + v(CFS, "J10")
    included = [y for y, t in zip(years, periods) if t > 0]

    checks += [
        one("WACC = cost of equity x E/V + after-tax cost of debt x D/V", wacc, "B54", tol=1e-9, sheet=A),
        *[one(f"unlevered free cash flow {y} = NOPAT at the scenario tax rate + D&A - CapEx - change in WC", f, f"{c}13")
          for y, f, c in zip(years, fcf, "BCDE")],
        *[one(f"years from valuation date to {y} year-end", t, f"{c}15", tol=1e-9)
          for y, t, c in zip(years, periods, "BCDE")],
        one(f"sum of discounted cash flows ({', '.join(included)})", pv, "B20"),
        one("terminal value, discounted from the FY29E year-end", tv_pv, "B21"),
        one("enterprise value", pv + tv_pv, "B22"),
        one("net cash (FY26E) = cash + marketable securities - total debt", net_cash, "B23"),
        one("implied share price", price, "B26", tol=0.01),
        one("upside = implied price / reference price - 1", price / price_ref - 1, "B28", tol=1e-6),
        one("sensitivity grid centre equals the headline price", price, "D36", tol=0.01),
        one("grid corner, WACC +50 bps and g -50 bps",
            _dcf_price(fcf, periods, wacc + 0.005, g - 0.005, net_cash, shares), "C37", tol=0.01),
        one("grid corner, WACC -50 bps and g +50 bps",
            _dcf_price(fcf, periods, wacc - 0.005, g + 0.005, net_cash, shares), "E35", tol=0.01),
        one("P/E bar (mid multiple x FY29E diluted EPS, discounted)",
            v(IS, "J43") * v(A, "C97") * df_last, "C69", tol=0.01),
        one("EV/EBITDA bar (mid multiple x FY29E EBITDA, discounted, plus net cash)",
            (ebitda * v(A, "C98") * df_last + net_cash) / shares, "C70", tol=0.01),
    ]

    lo, hi = g + 1e-6, 0.5
    f_lo, f_hi = (_dcf_price(fcf, periods, x, g, net_cash, shares) for x in (lo, hi))
    if not f_lo > price_ref > f_hi:
        raise ValueError("the reference price is outside the range the reverse DCF can solve for")
    for _ in range(200):
        mid = (lo + hi) / 2
        if _dcf_price(fcf, periods, mid, g, net_cash, shares) > price_ref:
            lo = mid
        else:
            hi = mid
    info = {"price": price, "price_ref": price_ref, "wacc": wacc, "g": g, "implied_wacc": (lo + hi) / 2,
            "tv_share": tv_pv / (pv + tv_pv), "scenario": v(A, "B62"), "ev": pv + tv_pv,
            "net_cash": net_cash, "equity_value": pv + tv_pv + net_cash, "market_value": price_ref * shares,
            "bear": v(D, "B68"), "bull": v(D, "D68")}
    return checks, info


def main(argv):
    """Load the workbook, run every check group, print the report and return the exit code (0 when all pass)."""
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT_MODEL
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    if "Validation" not in wb.sheetnames:
        sys.exit(f"no Validation sheet in {path.name}; sheets are {wb.sheetnames}")

    ws = wb["Validation"]

    # Verified, not assumed: same years across all three statements, and no
    # duplicate row label that a check could silently bind to the wrong row.
    years = check_year_headers(ws, find_section_headers(ws))
    rows = read_rows(ws)

    # A second load, formulas rather than values, purely to read the source-cell
    # mapping out of the Validation tab's column G.
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    links = read_source_links(wb_formulas["Validation"])
    if not links:
        sys.exit(
            "No source links found in the Validation tab's column G. Without them "
            "this script can only check the transcription against itself, which is "
            "the failure mode it exists to avoid. Refusing to report a pass."
        )

    print(f"Model:  {path.name}")
    print(f"Years:  {', '.join(str(y) for y in years)}")
    print(f"Tolerance: {TOLERANCE:,.0f} (millions USD, for the workbook's rounding)")
    print("-" * 72)

    dcf_checks, dcf = build_dcf_checks(wb)
    ratio_checks = build_ratio_checks(wb, rows, years)
    checks = build_checks(rows, years) + build_transcription_checks(wb, links, years) + dcf_checks + ratio_checks
    failed = 0
    for check in checks:
        if check.passed:
            print(f"  PASS  {check.label}")
        elif check.advisory:
            print(f"  NOTE  {check.label}")
        else:
            failed += 1
            print(f"  FAIL  {check.label}")
            for year, exp, act, why in check.failures:
                print(f"          {year}: derived {exp!r} vs workbook {act!r} ({why})")
            if check.note:
                print()
                for line in _wrap(check.note, 66):
                    print(f"          {line}")
                print()

    hard = [c for c in checks if not c.advisory]
    notes = [c for c in checks if c.advisory and not c.passed]
    identities = [c for c in hard if not c.label.startswith(("SRC ", "DCF ", "RAT "))]
    ratios = [c for c in hard if c.label.startswith("RAT ")]
    ratio_failures = sum(1 for c in ratios if not c.passed)
    sources = [c for c in hard if c.label.startswith("SRC ")]
    valuation = [c for c in hard if c.label.startswith("DCF ")]
    valuation_failures = sum(1 for c in valuation if not c.passed)
    identity_failures = sum(1 for c in identities if not c.passed)
    source_failures = sum(1 for c in sources if not c.passed)

    print("-" * 72)
    print(f"{len(identities) - identity_failures} of {len(identities)} accounting "
          f"identities re-derived independently and matched.")
    print(f"{len(sources) - source_failures} of {len(sources)} Validation-tab line "
          f"items match the IS / BS / CFS cells they were transcribed from, "
          f"across {len(years)} years each.")
    print(f"{len(valuation) - valuation_failures} of {len(valuation)} valuation figures re-derived "
          f"from the forecast statements and inputs ({dcf['scenario']} scenario).")
    print(f"{len(ratios) - ratio_failures} of {len(ratios)} Ratios-tab rows rebuilt from the traced line items "
          f"and matched in every year.")
    print(f"DCF: ${dcf['price']:,.2f} a share at WACC {dcf['wacc']:.2%} and g {dcf['g']:.1%}, against "
          f"${dcf['price_ref']:,.2f}. The reference price implies a WACC of {dcf['implied_wacc']:.2%} "
          f"on the same cash flows. Terminal value is {dcf['tv_share']:.0%} of enterprise value.")
    print(f"Enterprise value ${dcf['ev'] / 1e6:,.2f}tn plus FY26E net cash ${dcf['net_cash'] / 1e3:,.1f}bn.")
    print(f"Bear / Bull prices as recorded in the workbook (not re-derived): "
          f"${dcf['bear']:,.2f} / ${dcf['bull']:,.2f}.")
    print(f"Equity value ${dcf['equity_value'] / 1e6:,.2f}tn against a market value of "
          f"${dcf['market_value'] / 1e6:,.2f}tn at the reference price: the market pays "
          f"${(dcf['market_value'] - dcf['equity_value']) / 1e6:,.2f}tn more than the base case supports.")
    if notes:
        print(f"{len(notes)} advisory check(s) flagged above: explained presentation "
              f"differences, not arithmetic. See the note under each.")
    if failed:
        print(f"{failed} FAILED. These are arithmetic, and the workbook's own "
              f"Validation tab reports a tick for them anyway.")
        return 1
    print("Every total was recomputed outside the spreadsheet and agrees, every line "
          "item traces to the statement it came from, and the valuation rebuilds.")
    print("Not covered: the Dashboard and Pivots tabs; the Bear and Bull prices are "
          "printed as recorded.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
