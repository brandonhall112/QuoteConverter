from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
import json

DEFAULT_ALLOWED_REPS = [
    "Reid Kincaid",
    "LuWanna Morris",
    "Brent Schrader",
    "Eric Simpson",
    "Tami Knoell",
    "Alisha Teslow",
    "Darryl Overstreet",
]

QUOTE_SYNONYMS = {
    "quote_number": ["Quote", "Quote #", "Quote Number", "Quote No", "QuoteNum"],
    "sales_order": [
        "Sales Order",
        "Sales Order #",
        "Sales Order Number",
        "Order Number",
        "SO",
        "SO #",
    ],
    "customer": ["Customer", "Customer Name", "Cust. Name", "Account", "Bill To Name"],
    "quote_amount": ["Quote Amount", "Amount", "Total", "Quoted Total"],
    "date_quoted": ["Date Quoted", "Quote Date", "Entry Date", "Quoted Date"],
    "entry_person_name": [
        "Entry Person Name",
        "Primary Sales Rep",
        "Sales Rep Name",
        "Entry Person",
        "Created By",
    ],
    "rev": ["Rev", "Revision", "Quote Rev", "Quote Revision"],
}

ORDER_SYNONYMS = {
    "customer": ["Customer", "Customer Name", "Cust. Name", "Account", "Bill To Name"],
    "net": ["Net", "Net Price", "Net Amount", "Net USD", "NetValue", "Ext Net"],
    "rev": ["Rev", "Revision", "Order Rev"],
    "open": ["Open", "Is Open", "Open?"],
    "void": ["Void", "Voided"],
    "order_id": [
        "Order",
        "Order Number",
        "SO",
        "Sales Order",
        "Document",
        "Document Number",
        "Order No",
    ],
}


class FollowupError(Exception):
    """Expected domain error to display cleanly in CLI."""


@dataclass
class ColumnMap:
    quotes: dict[str, str] = field(default_factory=dict)
    orders: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_json(cls, path: str | Path | None) -> "ColumnMap":
        if not path:
            return cls()
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls(quotes=raw.get("quotes", {}), orders=raw.get("orders", {}))


@dataclass
class RunConfig:
    quotes_path: Path
    orders_path: Path
    out_path: Path
    floor: float = 1500.0
    tolerance: float = 1.0
    relative_tolerance: float = 0.05
    sheet_quotes: str | None = None
    sheet_orders: str | None = None
    reps: list[str] = field(default_factory=lambda: DEFAULT_ALLOWED_REPS.copy())
    debug: bool = False
    fuzzy: bool = False
    fuzzy_threshold: int = 90
    column_map: ColumnMap = field(default_factory=ColumnMap)


def load_reps(reps: list[str] | None, reps_config: str | None) -> list[str]:
    if reps:
        return reps
    if reps_config:
        parsed: Any = json.loads(Path(reps_config).read_text(encoding="utf-8"))
        if not isinstance(parsed, list) or not all(isinstance(x, str) for x in parsed):
            raise FollowupError("reps.json must be a JSON array of names.")
        return parsed
    return DEFAULT_ALLOWED_REPS.copy()
