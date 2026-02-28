from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import RunConfig
from .io_excel import normalize_customer, parse_money

OUTPUT_COLUMNS = ["Quote", "Customer", "Quote Amount", "Date Quoted", "Entry Person Name", "Won by Follow Up?"]


@dataclass
class MatchResult:
    followups: pd.DataFrame
    meta: pd.DataFrame
    debug: pd.DataFrame | None


def _normalize_order_id(value: object) -> str | None:
    if pd.isna(value):
        return None
    token = str(value).strip()
    return token or None


def _money_match(order_total: float, quote_amount: float, cfg: RunConfig) -> bool:
    diff = abs(order_total - quote_amount)
    if diff <= 0.005:
        return True
    relative_limit = abs(quote_amount) * cfg.relative_tolerance
    effective_tolerance = max(cfg.tolerance, relative_limit)
    return diff <= effective_tolerance


def _prep_quotes(quotes: pd.DataFrame, qmap: dict[str, str], cfg: RunConfig) -> pd.DataFrame:
    q = quotes.copy()
    q["Quote"] = q[qmap["quote_number"]]
    q["Customer"] = q[qmap["customer"]]
    q["Quote Amount"] = q[qmap["quote_amount"]].map(parse_money)
    q["Date Quoted"] = q[qmap["date_quoted"]]
    q["Entry Person Name"] = q[qmap["entry_person_name"]]
    q["CustKey"] = q["Customer"].map(normalize_customer)
    q["Won by Follow Up?"] = False

    q = q[q["Quote Amount"].notna()]
    q = q[q["Quote Amount"] > cfg.floor]
    q = q[q["Entry Person Name"].isin(cfg.reps)]
    return q


def _prep_order_totals(orders: pd.DataFrame, omap: dict[str, str]) -> pd.DataFrame:
    o = orders.copy()
    o["Customer"] = o[omap["customer"]]
    o["CustKey"] = o["Customer"].map(normalize_customer)
    o["Net"] = o[omap["net"]].map(parse_money)
    o = o[o["Net"].notna()]

    if "order_id" not in omap:
        totals = o.groupby(["CustKey"], dropna=False).agg(OrderTotal=("Net", "sum")).reset_index()
        return totals

    o["OrderId"] = o[omap["order_id"]].map(_normalize_order_id)
    totals = o.groupby(["CustKey", "OrderId"], dropna=False).agg(OrderTotal=("Net", "sum")).reset_index()
    return totals[["CustKey", "OrderTotal"]]


def _build_customer_index(order_totals: pd.DataFrame) -> dict[str, list[float]]:
    idx: dict[str, list[float]] = {}
    for _, row in order_totals.iterrows():
        idx.setdefault(str(row["CustKey"]), []).append(float(row["OrderTotal"]))
    return idx


def _quote_is_matched(quote_row: pd.Series, index: dict[str, list[float]], cfg: RunConfig) -> bool:
    amounts = index.get(str(quote_row["CustKey"]), [])
    qamt = float(quote_row["Quote Amount"])
    return any(_money_match(oa, qamt, cfg) for oa in amounts)


def _dedupe_sort(df: pd.DataFrame) -> pd.DataFrame:
    out = df.drop_duplicates(subset=OUTPUT_COLUMNS, keep="first")
    return out.sort_values(by=["Entry Person Name", "Customer", "Quote Amount"], ascending=[True, True, False])


def run_matching(quotes: pd.DataFrame, orders: pd.DataFrame, qmap: dict[str, str], omap: dict[str, str], cfg: RunConfig) -> MatchResult:
    q = _prep_quotes(quotes, qmap, cfg)
    order_totals = _prep_order_totals(orders, omap)

    cust_index = _build_customer_index(order_totals)
    q["Matched"] = q.apply(lambda r: _quote_is_matched(r, cust_index, cfg), axis=1)

    followups = _dedupe_sort(q[~q["Matched"]].copy())[OUTPUT_COLUMNS]

    meta_rows = [
        ("quotes_total_filtered", len(q)),
        ("followups", len(followups)),
        ("floor", cfg.floor),
        ("tolerance", cfg.tolerance),
        ("relative_tolerance", cfg.relative_tolerance),
        ("reps_count", len(cfg.reps)),
        ("quotes_mapping", str(qmap)),
        ("orders_mapping", str(omap)),
        ("notes", "Only Option B logic is used (customer + grouped order totals + tolerance)."),
    ]
    meta = pd.DataFrame(meta_rows, columns=["Metric", "Value"])

    debug = None
    if cfg.debug:
        debug = q[
            [
                "Quote",
                "Customer",
                "Quote Amount",
                "Date Quoted",
                "Entry Person Name",
                "Matched",
            ]
        ].copy()

    return MatchResult(followups=followups, meta=meta, debug=debug)
