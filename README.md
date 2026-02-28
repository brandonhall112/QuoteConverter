# Follow-Up Quote Finder

This tool compares a **Quote Summary** Excel export against an **Order Log** Excel export and produces follow-up call lists.

It outputs one Excel workbook with:
- **Option A (Rev Match)**: follow-ups where customer + amount + Rev must match (uses absolute/relative tolerance; falls back if Rev missing)
- **Option B (No Rev Match)**: follow-ups where customer + amount match (uses absolute/relative tolerance; Rev ignored)
- **Option C (Open Matched)**: sanity list of quotes that match an **Open** order (customer + amount tolerance match)

The output workbook contains **values only** (no formulas/macros) to avoid Excel security prompts.

---

## What you need each run

1. **Quote Summary (.xlsx)** for the date range
2. **Order Log (.xlsx)** for the same date range
3. (Optional) Follow-up Summary template workbook (.xlsx) if you want to preserve summary formulas/layout
4. Run the tool and send the output workbook to your team

---

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

Or installed command:

```bash
followup_quotes_ui
```

Then:
1. Browse and select **Quote Summary** file
2. Browse and select **Order Log** file
3. Choose output `.xlsx` path
4. Click **Generate Follow-Up Workbook**

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
- `--template "Followup_Template.xlsx"` (optional; writes results into existing template sheets and preserves other formula tabs)
- `--debug`
- `--fuzzy --fuzzy-threshold 90` (accepted; default matching remains normalized-exact)


## Template output behavior

When `--template` is provided:
- the template workbook is copied to the output path
- data sheets (`Option A (Rev Match)`, `Option B (No Rev Match)`, `Option C (Open Matched)`, `_Meta`) are refreshed in-place
- non-output sheets (for example a summary tab with formulas/charts) are preserved

This lets you keep your summary formulas/macros/layout while updating follow-up data from the latest run.

## Mapping override format

```json
{
  "quotes": {
    "quote_number": "Quote #",
    "customer": "Customer Name",
    "quote_amount": "Amount",
    "date_quoted": "Quote Date",
    "entry_person_name": "Created By",
    "rev": "Rev"
  },
  "orders": {
    "customer": "Customer",
    "net": "Net Amount",
    "order_id": "Order Number",
    "rev": "Revision",
    "open": "Open"
  }
}
```

## Notes

- Order logs are often line-level; this tool groups lines to order totals by Sales Order/Order ID when available.
- Quote numbers are not expected to equal order numbers; matching compares customer + order-level totals from the order log against quote totals.
- If Rev is missing in either file, Option A automatically falls back to Option B and records this in `_Meta`.
- If an Open column is missing, Option C is produced as an empty sheet and `_Meta` notes it.

- UI includes optional template and icon selectors, plus a refreshed modernized layout.

## CI build file (`.yml`)

This repository includes a GitHub Actions workflow at:
- `.github/workflows/build.yml`

What it does:
- Triggers on manual run (`workflow_dispatch`), `push`, and `pull_request`
- Runs tests on Ubuntu (Python 3.11 and 3.12)
- Builds a **Windows executable** for the desktop UI using PyInstaller
- Uploads the `.exe` as an Actions artifact named `FollowUpQuoteFinder-windows`

### Getting the executable from Actions

1. Open GitHub **Actions** tab
2. Run **Build and Test** workflow (or use an existing successful run)
3. Open run details
4. Download artifact: `FollowUpQuoteFinder-windows`
5. Extract and use `FollowUpQuoteFinder.exe`

### Why “Run workflow” can be missing

If you do not see a **Run workflow** button in GitHub Actions:
- Make sure the workflow file is on the repository's default branch (usually `main`).
- Make sure `workflow_dispatch:` exists in the workflow `on:` block.
- Ensure Actions are enabled for the repository/org.
