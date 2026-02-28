# Follow-Up Quote Finder

This tool compares a **Quote Summary** Excel export against an **Order Log** Excel export and produces one follow-up list of **unconverted quotes**.

It outputs one Excel workbook with:
- **Follow-Up**: quotes needing follow-up based on customer + amount closeness match against grouped sales-order totals
  - includes `Won by Follow Up?` column (boolean) so your summary checkbox/count logic keeps working
- **Per-rep tabs**: one sheet per `Entry Person Name` containing only that rep's follow-up lines
- **_Meta**: run/config diagnostics

The output workbook contains values in data tabs, and can preserve template formulas/layout when a template file is provided.

---

## What you need each run

1. **Quote Summary (.xlsx)** for the date range
2. **Order Log (.xlsx)** for the same date range
3. Put your template/icon files in `assets/` so they are auto-detected
4. Run the tool and send the output workbook to your team

---


## Assets folder

Create/use this folder in repo root:

- `assets/app.ico` (primary app/exe icon)
- `assets/followup.ico` (fallback icon)
- `assets/Parts Follow Up Template.xlsx` (default/preferred template)

The app auto-detects these assets by default (you can still override template path).

## Install

### Windows (recommended)
1. Install Python 3.11+ from python.org
2. Open Command Prompt (cmd)
3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage (CLI)

```bash
python -m followup_quotes.cli ^
  --quotes "Quote Summary.xlsx" ^
  --orders "Order Log.xlsx" ^
  --out "FollowUp_Output.xlsx"
```

Equivalent installed command:

```bash
followup_quotes --quotes "Quote Summary.xlsx" --orders "Order Log.xlsx" --out "FollowUp_Output.xlsx"
```

## Usage (Desktop UI)

Run:

```bash
python -m followup_quotes.ui
```

## CLI

Required:
- `--quotes <path>`
- `--orders <path>`
- `--out <path>`

Optional:
- `--floor 1500`
- `--tolerance 1`
- `--relative-tolerance 0.05` (5% default; matching uses max of absolute and relative tolerance)
- `--sheet-quotes "SheetName"`
- `--sheet-orders "SheetName"`
- `--reps "Name1" "Name2" ...`
- `--reps-config reps.json`
- `--column-map mapping.json`
- `--template "Followup_Template.xlsx"` (optional override; if omitted, app auto-detects templates and prefers `assets/Parts Follow Up Template.xlsx` (checks assets in current folder and executable folder first))
- `--debug`

## Template output behavior

When `--template` is provided:
- the template workbook is copied to the output path (auto-detected if not explicitly provided)
- output sheets (`Follow-Up`, rep tabs, `_Meta`) are refreshed/created in-place
- existing Excel table objects on template sheets are updated in-place (header/rows/ref), which avoids table repair prompts on open
- non-output sheets (for example a summary tab with formulas/charts) are preserved

## Mapping override format

```json
{
  "quotes": {
    "quote_number": "Quote #",
    "customer": "Customer Name",
    "quote_amount": "Amount",
    "date_quoted": "Quote Date",
    "entry_person_name": "Created By"
  },
  "orders": {
    "customer": "Customer",
    "net": "Net Amount",
    "order_id": "Order Number"
  }
}
```

## Notes

- Matching is **Option B only**: customer + grouped order totals + tolerance.
- Rev matching is not used.
- Quote numbers are not expected to equal order numbers; matching compares customer + order-level totals from the order log against quote totals.
- UI automatically applies an app icon when `assets/app.ico` (or `assets/followup.ico`) exists, including packaged executable locations.
- UI can auto-detect the template path and still allows override if needed.

## CI build file (`.yml`)

This repository includes a GitHub Actions workflow at:
- `.github/workflows/build.yml`

What it does:
- Triggers **manually only** (`workflow_dispatch`)
- Runs tests on Ubuntu (Python 3.11 and 3.12)
- Builds a **Windows executable** for the desktop UI using PyInstaller
- Uploads the `.exe` as an Actions artifact named `FollowUpQuoteFinder-windows`
