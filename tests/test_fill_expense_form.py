"""Run with: python -m pytest tests/"""
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import fill_expense_form as fef  # noqa: E402

SAMPLE = ROOT / "examples" / "sample_expenses.xlsx"


def test_parse_money():
    assert fef.parse_money("$1,308.29") == Decimal("1308.29")
    assert fef.parse_money("($5.00)") == Decimal("-5.00")
    assert fef.parse_money(12.5) == Decimal("12.5")
    assert fef.parse_money(None) == 0


def test_summary_of_sample_file():
    rows = fef.summarize(fef.load_expenses(SAMPLE))
    by_cat = {r["category"]: r for r in rows}

    assert "Miscellaneous" not in by_cat  # $0 rows are skipped
    assert [r["category"] for r in rows] == [
        "Conference Registration",
        "Airfare",
        "Ground Transportation",
        "Individual Meals",
        "Hotel-Individual",
    ]
    assert by_cat["Individual Meals"]["total"] == Decimal("46.10")
    assert by_cat["Individual Meals"]["count"] == 2
    assert by_cat["Ground Transportation"]["total"] == Decimal("57.75")
    assert sum(r["total"] for r in rows) == Decimal("1291.55")


def test_cli_writes_filled_pdf(tmp_path):
    out = tmp_path / "form.pdf"
    subprocess.run(
        [sys.executable, str(ROOT / "fill_expense_form.py"), str(SAMPLE),
         "--name", "Alex Example", "--advisor", "Prof. Sample",
         "--chartstring", "C0000-TEST", "-o", str(out)],
        check=True,
    )
    text = PdfReader(out).pages[0].extract_text()
    for expected in ["Alex Example", "Prof. Sample", "C0000-TEST", "$1,291.55",
                     "Airfare (1 transaction)", "Individual Meals (2 transactions)",
                     "03/10/26 - 03/14/26", "$612.30"]:
        assert expected in text
