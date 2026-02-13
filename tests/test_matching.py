from pathlib import Path

import pandas as pd

from followup_quotes.config import RunConfig
from followup_quotes.matching import run_matching


def test_option_a_b_c_and_grouping():
    quotes = pd.DataFrame(
        {
            "Quote #": ["Q1", "Q2", "Q3", "Q4"],
            "Customer Name": ["Acme, Inc.", "Acme Inc", "Beta LLC", "Gamma"],
            "Amount": [2000, 2500, 3000, 4000],
            "Quote Date": ["2024-01-01"] * 4,
            "Created By": ["Reid Kincaid"] * 4,
            "Rev": ["A", "B", "X", "X"],
        }
    )
    orders = pd.DataFrame(
        {
            "Order Number": [10, 10, 20, 30],
            "Customer": ["ACME INC", "ACME INC", "BETA LLC", "GAMMA"],
            "Net Amount": [1000, 1000, 3000, 4000],
            "Revision": ["A", "A", "X", "Y"],
            "Open": ["YES", "YES", "NO", "TRUE"],
        }
    )
    qmap = {
        "quote_number": "Quote #",
        "customer": "Customer Name",
        "quote_amount": "Amount",
        "date_quoted": "Quote Date",
        "entry_person_name": "Created By",
        "rev": "Rev",
    }
    omap = {
        "order_id": "Order Number",
        "customer": "Customer",
        "net": "Net Amount",
        "rev": "Revision",
        "open": "Open",
    }
    cfg = RunConfig(
        quotes_path=Path("q.xlsx"),
        orders_path=Path("o.xlsx"),
        out_path=Path("x.xlsx"),
        reps=["Reid Kincaid"],
    )

    out = run_matching(quotes, orders, qmap, omap, cfg)

    assert set(out.option_b["Quote"]) == {"Q2"}
    assert set(out.option_a["Quote"]) == {"Q2", "Q4"}
    assert set(out.option_c["Quote"]) == {"Q1", "Q3", "Q4"}


def test_rev_fallback_when_missing():
    quotes = pd.DataFrame(
        {
            "Quote #": ["Q1"],
            "Customer": ["Acme"],
            "Amount": [2000],
            "Date Quoted": ["2024-01-01"],
            "Entry Person Name": ["Reid Kincaid"],
        }
    )
    orders = pd.DataFrame({"Customer": ["ACME"], "Net": [2000]})
    qmap = {
        "quote_number": "Quote #",
        "customer": "Customer",
        "quote_amount": "Amount",
        "date_quoted": "Date Quoted",
        "entry_person_name": "Entry Person Name",
    }
    omap = {"customer": "Customer", "net": "Net"}
    cfg = RunConfig(
        quotes_path=Path("q.xlsx"),
        orders_path=Path("o.xlsx"),
        out_path=Path("x.xlsx"),
        reps=["Reid Kincaid"],
    )

    out = run_matching(quotes, orders, qmap, omap, cfg)
    assert out.option_a.equals(out.option_b)
    assert "fell back" in str(out.meta[out.meta["Metric"] == "notes"]["Value"].iloc[0])
