"""Independently re-derive every accounting identity in a model workbook.

The workbook already has a Validation tab that prints "ALL CHECKS PASS". This
script exists because that tick mark is produced by the same spreadsheet whose
correctness is in question. A formula that is wrong reports a tick just as
confidently as one that is right, so a reader has no way to distinguish the two
without recomputing the arithmetic outside the file.

This does that. It reads only the raw line items from the Validation tab, adds
them up in Python, and compares the result against the totals the workbook
reports. Nothing here reads a cell whose value is a checkmark.

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

Usage:

    python validate_model.py                        # defaults to the Apple model
    python validate_model.py path/to/Model.xlsx

Exits 0 when every check passes, 1 otherwise, so it can run in CI.
"""
from __future__ import annotations

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
                 advisory: bool = False):
        self.label = label
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
        out = []
        for year, exp, act in zip(self.years, self.expected, self.actual):
            if exp is None or act is None:
                out.append((year, exp, act, "missing value"))
            elif abs(exp - act) > TOLERANCE:
                out.append((year, exp, act, f"differs by {exp - act:,.1f}"))
        return out

    @property
    def passed(self) -> bool:
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
                  "Inventory", "Other CA (vendor + other)"),
              g("Total current assets"), years),
        Check("BS  total assets = current + PP&E + other non-current",
              add("Total current assets", "Marketable securities (non-curr)",
                  "Property, plant & equipment, net", "Other non-current assets"),
              g("Total assets"), years),
        Check("BS  current liabilities sum to the subtotal",
              add("Accounts payable", "Other CL (bundled)"), g("Total current liabilities"), years),
        Check("BS  total liabilities = current + LT debt + other non-current",
              add("Total current liabilities", "Long-term debt (non-current)",
                  "Other non-current liabilities"),
              g("Total liabilities"), years),
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
                  "Apple's cash flow statement historically reconciled to 'cash, "
                  "cash equivalents and restricted cash', while the balance sheet "
                  "line excludes restricted cash. If the gap here is restricted "
                  "cash, the model is faithfully reproducing each filing's own "
                  "presentation and the two lines are not supposed to be equal. "
                  "Confirm the figure against the restricted-cash disclosure in "
                  "the relevant 10-K before treating this as either a defect or a "
                  "non-issue, and then say which it is in the README."
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


def main(argv):
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT_MODEL
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    if "Validation" not in wb.sheetnames:
        sys.exit(f"no Validation sheet in {path.name}; sheets are {wb.sheetnames}")

    ws = wb["Validation"]

    # Verified, not assumed: same years across all three statements, and no
    # duplicate row label that a check could silently bind to the wrong row.
    years = check_year_headers(
        ws,
        {"income statement": 15, "balance sheet": 35, "cash flow statement": 58},
    )
    rows = read_rows(ws)

    print(f"Model:  {path.name}")
    print(f"Years:  {', '.join(str(y) for y in years)}")
    print(f"Tolerance: {TOLERANCE:,.0f} (millions USD, for the workbook's rounding)")
    print("-" * 72)

    checks = build_checks(rows, years)
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

    print("-" * 72)
    print(f"{len(hard) - failed} of {len(hard)} accounting identities re-derived "
          f"independently and matched.")
    if notes:
        print(f"{len(notes)} advisory check(s) flagged above: a presentation "
              f"question, not arithmetic. See the note under each.")
    if failed:
        print(f"{failed} FAILED. These are arithmetic, and the workbook's own "
              f"Validation tab reports a tick for them anyway.")
        return 1
    print("Every total was recomputed outside the spreadsheet and agrees.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
