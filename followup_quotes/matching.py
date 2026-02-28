from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import RunConfig
from .io_excel import normalize_customer, parse_money, parse_truthy

OUTPUT_COLUMNS = ["Quote", "Customer", "Quote Amount", "Date Quoted", "Entry Person Name"]


@dataclass
class MatchResult:
    option_a: pd.DataFrame
    option_b: pd.DataFrame
    option_c: pd.DataFrame
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
    q["Rev"] = q[qmap["rev"]] if "rev" in qmap else None
    q["CustKey"] = q["Customer"].map(normalize_customer)

    q = q[q["Quote Amount"].notna()]
    q = q[q["Quote Amount"] > cfg.floor]
    q = q[q["Entry Person Name"].isin(cfg.reps)]
    return q


def _prep_order_candidates(orders: pd.DataFrame, omap: dict[str, str]) -> pd.DataFrame:
    o = orders.copy()
    o["Customer"] = o[omap["customer"]]
    o["CustKey"] = o["Customer"].map(normalize_customer)
    o["Net"] = o[omap["net"]].map(parse_money)
    o = o[o["Net"].notna()]
    o["Open"] = o[omap["open"]].map(parse_truthy) if "open" in omap else False
    o["Rev"] = o[omap["rev"]] if "rev" in omap else None

    if "order_id" not in omap:
        out = o[["CustKey", "Net", "Open", "Rev"]].rename(columns={"Net": "OrderTotal"})
        out["OrderId"] = None
        return out

    o["OrderId"] = o[omap["order_id"]].map(_normalize_order_id)
    grouped = (
        o.groupby(["OrderId", "CustKey"], dropna=False)
        .agg(OrderTotal=("Net", "sum"), OpenOrder=("Open", "max"))
        .reset_index()
    )

    rev_sets = (
        o.groupby(["OrderId", "CustKey"], dropna=False)["Rev"]
        .apply(lambda s: {str(v).strip() for v in s if pd.notna(v) and str(v).strip() != ""})
        .reset_index(name="RevSet")
    )
    merged = grouped.merge(rev_sets, on=["OrderId", "CustKey"], how="left")

    rows = []
    for _, row in merged.iterrows():
        revset = row["RevSet"] or set()
        if not revset:
            rows.append({**row.to_dict(), "Rev": None, "Open": bool(row["OpenOrder"])})
        else:
            for rev in revset:
                rows.append({**row.to_dict(), "Rev": rev, "Open": bool(row["OpenOrder"])})

    return pd.DataFrame(rows)[["OrderId", "CustKey", "OrderTotal", "Open", "Rev"]]


def _build_index(order_candidates: pd.DataFrame, with_rev: bool) -> dict[tuple, list[float]]:
    idx: dict[tuple, list[float]] = {}
    for _, row in order_candidates.iterrows():
        if with_rev:
            key = (row["CustKey"], None if pd.isna(row["Rev"]) else str(row["Rev"]).strip())
        else:
            key = (row["CustKey"],)
        idx.setdefault(key, []).append(float(row["OrderTotal"]))
    return idx


def _quote_is_matched(
    quote_row: pd.Series,
    index: dict[tuple, list[float]],
    cfg: RunConfig,
    with_rev: bool,
) -> bool:
    if with_rev:
        rev = quote_row.get("Rev")
        key = (quote_row["CustKey"], None if pd.isna(rev) else str(rev).strip())
    else:
        key = (quote_row["CustKey"],)

    amounts = index.get(key, [])
    qamt = float(quote_row["Quote Amount"])
    return any(_money_match(oa, qamt, cfg) for oa in amounts)


def _dedupe_sort(df: pd.DataFrame, include_rev: bool) -> pd.DataFrame:
    dedupe_cols = OUTPUT_COLUMNS.copy()
    if include_rev and "Rev" in df.columns:
        dedupe_cols.append("Rev")
    out = df.drop_duplicates(subset=dedupe_cols, keep="first")
    return out.sort_values(by=["Entry Person Name", "Customer", "Quote Amount"], ascending=[True, True, False])


def run_matching(quotes: pd.DataFrame, orders: pd.DataFrame, qmap: dict[str, str], omap: dict[str, str], cfg: RunConfig) -> MatchResult:
    q = _prep_quotes(quotes, qmap, cfg)
    order_candidates = _prep_order_candidates(orders, omap)

    b_index = _build_index(order_candidates, with_rev=False)
    q["MatchedB"] = q.apply(lambda r: _quote_is_matched(r, b_index, cfg, with_rev=False), axis=1)

    rev_available = "rev" in qmap and "rev" in omap
    fallback_note = None
    if rev_available:
        a_index = _build_index(order_candidates, with_rev=True)
        q["MatchedA"] = q.apply(lambda r: _quote_is_matched(r, a_index, cfg, with_rev=True), axis=1)
    else:
        q["MatchedA"] = q["MatchedB"]
        fallback_note = "Rev not available; Option A fell back to Option B"

    if "open" in omap:
        open_index = _build_index(order_candidates[order_candidates["Open"]], with_rev=False)
        q["MatchedOpen"] = q.apply(lambda r: _quote_is_matched(r, open_index, cfg, with_rev=False), axis=1)
        open_note = ""
    else:
        q["MatchedOpen"] = False
        open_note = "Open column not found; Option C produced as empty sheet"

    option_a = _dedupe_sort(q[~q["MatchedA"]].copy(), include_rev="rev" in qmap)[OUTPUT_COLUMNS]
    option_b = _dedupe_sort(q[~q["MatchedB"]].copy(), include_rev="rev" in qmap)[OUTPUT_COLUMNS]
    option_c = _dedupe_sort(q[q["MatchedOpen"]].copy(), include_rev="rev" in qmap)[OUTPUT_COLUMNS]

    meta_rows = [
        ("quotes_total_filtered", len(q)),
        ("option_a_followups", len(option_a)),
        ("option_b_followups", len(option_b)),
        ("option_c_open_matched", len(option_c)),
        ("floor", cfg.floor),
        ("tolerance", cfg.tolerance),
        ("relative_tolerance", cfg.relative_tolerance),
        ("reps_count", len(cfg.reps)),
        ("rev_available_both_files", rev_available),
        ("quotes_mapping", str(qmap)),
        ("orders_mapping", str(omap)),
        ("notes", " | ".join(x for x in [fallback_note, open_note] if x)),
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
                "Rev",
                "MatchedA",
                "MatchedB",
                "MatchedOpen",
            ]
        ].copy()

    return MatchResult(option_a=option_a, option_b=option_b, option_c=option_c, meta=meta, debug=debug)
