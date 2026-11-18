# Expense Report Form Filler

Fills in the EEB **Lab and Travel Purchase/Reimbursement Request Form** (PDF) from a Concur expense export (Excel).

The script:
- fills in Name, Total amount being requested, Faculty Advisor, and Chartstring Code
- adds one table row per expense category, with its date range, number of transactions, and total
- uses the USD **Requested** column, so foreign-currency charges are already converted
- skips $0.00 lines

## Install

Requires Python 3.9+.

```bash
git clone <this repo>
cd expense_report_app
pip3 install -r requirements.txt
```

## Use

1. In Concur, export your report's expense list to Excel (`.xlsx`). It needs `Date`, `Expense Type`, and `Requested` columns. You can find the download button at the top right of the expense table when you have the draft expense report open. 
2. Run:

```bash
python3 fill_expense_form.py path/to/expenses.xlsx \
    --name "Your Name" \
    --advisor "Prof. Advisor" \
    --chartstring "XXXXX-XXXXX"
```

This writes `expenses_form.pdf` next to the Excel file and prints a summary to check against the PDF. Options:

| Option | Meaning |
|---|---|
| `-o out.pdf` | Choose the output file |
| `--template form.pdf` | Use a different copy of the blank form (default: the PDF in this folder) |

Get the faculty approval signature on the filled PDF, then attach it to your Concur report.

## Try it with the sample data

`examples/sample_expenses.xlsx` contains made-up expenses:

```bash
python3 fill_expense_form.py examples/sample_expenses.xlsx \
    --name "Alex Example" --advisor "Prof. Sample" --chartstring "C0000-TEST" \
    -o sample_form.pdf
```

Expected output:

```
Dates                Category                                  Amount
02/18/26             Conference Registration (1)              $250.00
02/20/26             Airfare (1)                              $612.30
03/10/26 - 03/14/26  Ground Transportation (2)                 $57.75
03/11/26 - 03/12/26  Individual Meals (2)                      $46.10
03/13/26             Hotel-Individual (1)                     $325.40
                     TOTAL                                  $1,291.55
```

## Tests

```bash
pip3 install pytest
python3 -m pytest tests/
```

## Notes

- The PDF has no fillable fields. Text is placed at fixed positions measured from this version of the form. If the department changes the form layout, update the coordinates near the top of `fill_expense_form.py`.
- The form has 16 table rows. The script stops with an error if there are more categories than that.
- `.gitignore` excludes `.xlsx` files (except the example) and filled `*_form.pdf` files so personal expense data isn't committed.
