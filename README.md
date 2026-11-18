# Expense Report Form Filler

Fills in the EEB **Lab and Travel Purchase/Reimbursement Request Form** (PDF) from a Concur expense export (Excel).

The script:
- fills in Name, Total amount being requested, Faculty Advisor, and Chartstring Code
- adds one table row per expense category, with its date range, number of transactions, and total
- uses the USD **Requested** column, so foreign-currency charges are already converted
- skips $0.00 lines

## Install

Requires Python 3.9+. If you've used the terminal before:

```bash
git clone <this repo>
cd expense_report_filler
pip3 install -r requirements.txt
```

### First time using the terminal?

1. **Install Python** from [python.org/downloads](https://www.python.org/downloads/). This also installs `pip`, which installs Python libraries.
   - **Windows:** in the installer, tick **"Add python.exe to PATH"**. In all commands below, type `python` instead of `python3` and `pip` instead of `pip3`.
2. **Get the code:** on the GitHub page, click **Code → Download ZIP**, then unzip it (for example into Downloads).
3. **Open the terminal:**
   - **Mac:** press `Cmd + Space`, type **Terminal**, press Enter.
   - **Windows:** open the Start menu, type **PowerShell**, press Enter.
4. **Go into the folder:** type `cd ` (with a space after it), drag the unzipped folder onto the terminal window, and press Enter.
5. **Install the libraries:**
   ```bash
   pip3 install -r requirements.txt
   ```
   On a new Mac, a pop-up may ask to install "command line developer tools". Click **Install**, wait for it to finish, then run the command again.

You only do this once. After that, open the terminal, `cd` into the folder (step 4), and follow **Use** below.

**Troubleshooting**
- `command not found: pip3`: run `python3 -m pip install -r requirements.txt` instead.
- `externally-managed-environment` error: you're using a Homebrew Python. Install Python from python.org instead.
- **File names with spaces** need quotes, like `"My Expenses.xlsx"`. Dragging a file onto the terminal fills in its path with the quoting done for you.

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
