#!/usr/bin/env python3
"""Fill the EEB Lab and Travel Purchase/Reimbursement Request Form from a Concur Excel export.

Expenses are summarized as one row per Expense Type (category), using the USD
"Requested" column. The PDF has no fillable fields, so text is drawn at fixed
positions and stamped onto the original form.

Example:
    python fill_expense_form.py SEGED_travel_2026.xlsx \
        --name "Jane Doe" --advisor "Prof. Smith" --chartstring "12345-ABC"
"""
import argparse
import io
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

DEFAULT_TEMPLATE = Path(__file__).with_name("Lab and Travel PurchaseReimbursement Request Form.pdf")

PAGE_HEIGHT = 792
FONT = "Helvetica"

# Layout measured from the template (PDF points, y measured from the TOP of the page).
# Header boxes: right-hand column spans x 278-540.
HEADER_X = 284
HEADER_ROWS = {  # (top, bottom) of each box
    "name": (169.6, 192.7),
    "total": (193.2, 215.8),
    "advisor": (216.3, 244.1),
    "chartstring": (244.6, 272.4),
}
# Table: 16 data rows between y=336.1 and y=597.7.
TABLE_ROW_EDGES = [336.1, 352.4, 368.7, 385.0, 401.5, 417.8, 434.1, 450.4,
                   466.9, 483.2, 499.5, 515.8, 532.3, 548.6, 564.9, 581.2, 597.7]
COL_DATE = (62.0, 161.5)
COL_DESC = (162.5, 459.8)
COL_AMOUNT = (460.8, 550.1)
CELL_PAD = 4


def parse_money(value) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).replace("$", "").replace(",", "").replace("\xa0", " ").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return Decimal(text)
    except InvalidOperation:
        raise ValueError(f"Can't read amount {value!r} (expected a USD value like $1,234.56)")


def parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Can't read date {value!r}")


def load_expenses(xlsx_path: Path):
    """Return a list of (date, category, usd_amount) from the Concur export."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip().lower() if h is not None else "" for h in next(rows)]

    def col(name):
        try:
            return header.index(name)
        except ValueError:
            sys.exit(f"Column '{name}' not found in {xlsx_path.name}. Columns are: {header}")

    i_date, i_type, i_req = col("date"), col("expense type"), col("requested")
    expenses = []
    for row in rows:
        if not row or all(c in (None, "") for c in row):
            continue
        amount = parse_money(row[i_req])
        if amount == 0:
            continue
        expenses.append((parse_date(row[i_date]), str(row[i_type]).strip(), amount))
    return expenses


def summarize(expenses):
    """Group by category. Returns rows sorted by earliest date."""
    groups = defaultdict(list)
    for d, cat, amt in expenses:
        groups[cat].append((d, amt))
    summary = []
    for cat, items in groups.items():
        dates = [d for d, _ in items]
        summary.append({
            "category": cat,
            "start": min(dates),
            "end": max(dates),
            "count": len(items),
            "total": sum((a for _, a in items), Decimal("0")),
        })
    summary.sort(key=lambda r: (r["start"], r["category"]))
    return summary


def fmt_money(amount: Decimal) -> str:
    return f"${amount:,.2f}"


def fmt_date_range(start: date, end: date) -> str:
    s = start.strftime("%m/%d/%y")
    return s if start == end else f"{s} - {end.strftime('%m/%d/%y')}"


def fit_font(text, max_width, size, min_size=6):
    while size > min_size and stringWidth(text, FONT, size) > max_width:
        size -= 0.5
    return size


def baseline(top, bottom, size):
    """Convert a top-down cell to a reportlab baseline that vertically centres text."""
    mid = (top + bottom) / 2
    return PAGE_HEIGHT - mid - size * 0.35


def build_overlay(name, advisor, chartstring, total, rows) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(612, PAGE_HEIGHT))

    header_values = {
        "name": name,
        "total": fmt_money(total),
        "advisor": advisor,
        "chartstring": chartstring,
    }
    for key, text in header_values.items():
        top, bottom = HEADER_ROWS[key]
        size = fit_font(text, 540 - HEADER_X - CELL_PAD, 11)
        c.setFont(FONT, size)
        c.drawString(HEADER_X, baseline(top, bottom, size), text)

    for i, row in enumerate(rows):
        top, bottom = TABLE_ROW_EDGES[i], TABLE_ROW_EDGES[i + 1]

        text = fmt_date_range(row["start"], row["end"])
        size = fit_font(text, COL_DATE[1] - COL_DATE[0] - 2 * CELL_PAD, 10)
        c.setFont(FONT, size)
        c.drawString(COL_DATE[0] + CELL_PAD, baseline(top, bottom, size), text)

        n = row["count"]
        text = f"{row['category']} ({n} transaction{'s' if n != 1 else ''})"
        size = fit_font(text, COL_DESC[1] - COL_DESC[0] - 2 * CELL_PAD, 10)
        c.setFont(FONT, size)
        c.drawString(COL_DESC[0] + CELL_PAD, baseline(top, bottom, size), text)

        text = fmt_money(row["total"])
        size = fit_font(text, COL_AMOUNT[1] - COL_AMOUNT[0] - 2 * CELL_PAD, 10)
        c.setFont(FONT, size)
        c.drawRightString(COL_AMOUNT[1] - CELL_PAD, baseline(top, bottom, size), text)

    c.save()
    return buf.getvalue()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", type=Path, help="Concur expense export (.xlsx)")
    ap.add_argument("--name", required=True, help="Your name")
    ap.add_argument("--advisor", required=True, help="Faculty advisor")
    ap.add_argument("--chartstring", required=True, help="Chartstring code")
    ap.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Blank form PDF")
    ap.add_argument("-o", "--output", type=Path, help="Output PDF (default: <xlsx name>_form.pdf)")
    args = ap.parse_args()

    expenses = load_expenses(args.xlsx)
    if not expenses:
        sys.exit("No non-zero expenses found.")
    rows = summarize(expenses)
    if len(rows) > len(TABLE_ROW_EDGES) - 1:
        sys.exit(f"{len(rows)} categories won't fit in the form's {len(TABLE_ROW_EDGES) - 1} rows.")
    total = sum((r["total"] for r in rows), Decimal("0"))

    overlay = PdfReader(io.BytesIO(build_overlay(args.name, args.advisor, args.chartstring, total, rows)))
    writer = PdfWriter(clone_from=args.template)
    writer.pages[0].merge_page(overlay.pages[0])

    out = args.output or args.xlsx.with_name(f"{args.xlsx.stem}_form.pdf")
    with open(out, "wb") as f:
        writer.write(f)

    print(f"{'Dates':<20} {'Category':<35} {'Amount':>12}")
    for r in rows:
        label = f"{r['category']} ({r['count']})"
        print(f"{fmt_date_range(r['start'], r['end']):<20} {label:<35} {fmt_money(r['total']):>12}")
    print(f"{'':<20} {'TOTAL':<35} {fmt_money(total):>12}")
    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
