from pathlib import Path

from openpyxl import Workbook, load_workbook
import pandas as pd

from followup_quotes.app import generate_followup_workbook, resolve_template_path
from followup_quotes.config import RunConfig
from followup_quotes.io_excel import write_output


def test_write_output_preserves_template_formula_sheet(tmp_path: Path):
    template = tmp_path / "template.xlsx"
    out = tmp_path / "out.xlsx"

    wb = Workbook()
    ws_followup = wb.active
    ws_followup.title = "Follow-Up"
    ws_followup["A1"] = "Quote"
    ws_followup["B1"] = "Customer"
    ws_summary = wb.create_sheet("Summary")
    ws_summary["A1"] = "Count"
    ws_summary["B1"] = "=COUNTA('Follow-Up'!A:A)-1"
    wb.save(template)

    sheets = {
        "Follow-Up": pd.DataFrame([{"Quote": "Q1", "Customer": "ACME"}, {"Quote": "Q2", "Customer": "BETA"}]),
        "_Meta": pd.DataFrame([{"Metric": "x", "Value": "y"}]),
    }

    write_output(out, sheets, template)

    out_wb = load_workbook(out)
    assert out_wb["Follow-Up"]["A2"].value == "Q1"
    assert out_wb["Follow-Up"]["A3"].value == "Q2"
    assert out_wb["Summary"]["B1"].value == "=COUNTA('Follow-Up'!A:A)-1"


def test_generate_followup_workbook_creates_per_rep_tabs(tmp_path: Path):
    quotes_path = tmp_path / "quotes.xlsx"
    orders_path = tmp_path / "orders.xlsx"
    out_path = tmp_path / "output.xlsx"

    pd.DataFrame(
        {
            "Quote #": ["Q1", "Q2"],
            "Customer": ["Acme", "Beta"],
            "Amount": [4000, 4100],
            "Date Quoted": ["2024-01-01", "2024-01-02"],
            "Entry Person Name": ["Reid Kincaid", "Eric Simpson"],
        }
    ).to_excel(quotes_path, index=False)

    pd.DataFrame(
        {
            "Order Number": [1],
            "Customer": ["ACME"],
            "Net Amount": [4000],
        }
    ).to_excel(orders_path, index=False)

    cfg = RunConfig(
        quotes_path=quotes_path,
        orders_path=orders_path,
        out_path=out_path,
        reps=["Reid Kincaid", "Eric Simpson"],
    )

    generate_followup_workbook(cfg)

    wb = load_workbook(out_path)
    assert "Follow-Up" in wb.sheetnames
    assert "Eric Simpson" in wb.sheetnames
    assert wb["Eric Simpson"]["A2"].value == "Q2"


def test_resolve_template_path_prefers_explicit(tmp_path: Path):
    explicit = tmp_path / "x.xlsx"
    explicit.write_bytes(b"dummy")
    assert resolve_template_path(explicit) == explicit
