# Follow-Up Quote Finder

This tool compares a **Quote Summary** Excel export against an **Order Log** Excel export and produces follow-up call lists.

It outputs one Excel workbook with:
- **Option A (Rev Match)**: follow-ups where customer + amount ±$1 + Rev must match (falls back if Rev missing)
- **Option B (No Rev Match)**: follow-ups where customer + amount ±$1 match (Rev ignored)
- **Option C (Open Matched)**: sanity list of quotes that match an **Open** order (customer + amount ±$1)

The output workbook contains **values only** (no formulas/macros) to avoid Excel security prompts.

---

## What you need each run

1. **Quote Summary (.xlsx)** for the date range
2. **Order Log (.xlsx)** for the same date range
3. Run the CLI command and send the output workbook to your team

---

## Install

### Windows (recommended)
1. Install Python 3.11+ from python.org
2. Open Command Prompt (cmd)
3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python -m followup_quotes.cli ^
  --quotes "Quote Summary.xlsx" ^
  --orders "Order Log.xlsx" ^
  --out "FollowUp_Output.xlsx"
```

```bash
python -m followup_quotes.cli ^
  --quotes "Quote Summary.xlsx" ^
  --orders "Order Log.xlsx" ^
  --out "FollowUp_Output.xlsx"
```

It is important to note that you can load an order log and quote summary on demand and generate an `.xlsx` workbook for team distribution with quotes to follow up.

## CLI

Required:
- `--quotes <path>`
- `--orders <path>`
- `--out <path>`

Optional:
- `--floor 1500`
- `--tolerance 1`
- `--sheet-quotes "SheetName"`
- `--sheet-orders "SheetName"`
- `--reps "Name1" "Name2" ...`
- `--reps-config reps.json`
- `--column-map mapping.json`
- `--debug`
- `--fuzzy --fuzzy-threshold 90` (accepted; default matching remains normalized-exact)

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

- Order logs are often line-level; this tool groups lines to order totals when an Order ID exists.
- If Rev is missing in either file, Option A automatically falls back to Option B and records this in `_Meta`.
- If an Open column is missing, Option C is produced as an empty sheet and `_Meta` notes it.
