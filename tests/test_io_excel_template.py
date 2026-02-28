from pathlib import Path

from openpyxl import Workbook, load_workbook
import pandas as pd

from followup_quotes.io_excel import write_output


def test_write_output_preserves_template_formula_sheet(tmp_path: Path):
    template = tmp_path / "template.xlsx"
    out = tmp_path / "out.xlsx"

    wb = Workbook()
    ws_a = wb.active
    ws_a.title = "Option B (No Rev Match)"
    ws_a["A1"] = "Quote"
    ws_a["B1"] = "Customer"
    ws_summary = wb.create_sheet("Summary")
    ws_summary["A1"] = "Count"
    ws_summary["B1"] = "=COUNTA('Option B (No Rev Match)'!A:A)-1"
    wb.save(template)

    sheets = {
        "Option B (No Rev Match)": pd.DataFrame(
            [{"Quote": "Q1", "Customer": "ACME"}, {"Quote": "Q2", "Customer": "BETA"}]
        ),
        "_Meta": pd.DataFrame([{"Metric": "x", "Value": "y"}]),
    }

    write_output(out, sheets, template)

    out_wb = load_workbook(out)
    assert out_wb["Option B (No Rev Match)"]["A2"].value == "Q1"
    assert out_wb["Option B (No Rev Match)"]["A3"].value == "Q2"
    assert out_wb["Summary"]["B1"].value == "=COUNTA('Option B (No Rev Match)'!A:A)-1"
