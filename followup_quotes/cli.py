from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .config import (
    ColumnMap,
    FollowupError,
    ORDER_SYNONYMS,
    QUOTE_SYNONYMS,
    RunConfig,
    load_reps,
)
from .io_excel import detect_columns, read_excel, write_output
from .matching import run_matching


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Generate follow-up quote workbook from Quote Summary and Order Log.")
    p.add_argument("--quotes", required=True, help="Path to Quote Summary xlsx")
    p.add_argument("--orders", required=True, help="Path to Order Log xlsx")
    p.add_argument("--out", required=True, help="Path for output xlsx")
    p.add_argument("--floor", type=float, default=1500)
    p.add_argument("--tolerance", type=float, default=1)
    p.add_argument("--sheet-quotes")
    p.add_argument("--sheet-orders")
    p.add_argument("--reps", nargs="*")
    p.add_argument("--reps-config")
    p.add_argument("--debug", action="store_true")
    p.add_argument("--column-map")
    p.add_argument("--fuzzy", action="store_true")
    p.add_argument("--fuzzy-threshold", type=int, default=90)
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        column_map = ColumnMap.from_json(args.column_map)
        reps = load_reps(args.reps, args.reps_config)
        cfg = RunConfig(
            quotes_path=Path(args.quotes),
            orders_path=Path(args.orders),
            out_path=Path(args.out),
            floor=args.floor,
            tolerance=args.tolerance,
            sheet_quotes=args.sheet_quotes,
            sheet_orders=args.sheet_orders,
            reps=reps,
            debug=args.debug,
            fuzzy=args.fuzzy,
            fuzzy_threshold=args.fuzzy_threshold,
            column_map=column_map,
        )

        quotes_df = read_excel(cfg.quotes_path, cfg.sheet_quotes)
        orders_df = read_excel(cfg.orders_path, cfg.sheet_orders)

        qdetect = detect_columns(
            quotes_df,
            QUOTE_SYNONYMS,
            required_fields={"quote_number", "customer", "quote_amount", "date_quoted", "entry_person_name"},
            overrides=cfg.column_map.quotes,
        )
        odetect = detect_columns(
            orders_df,
            ORDER_SYNONYMS,
            required_fields={"customer", "net"},
            overrides=cfg.column_map.orders,
            contains_rules={"open": "open", "void": "void"},
        )

        if cfg.fuzzy:
            qdetect.notes.append(f"fuzzy matching requested at threshold {cfg.fuzzy_threshold}, not needed for normalized-exact mode")

        result = run_matching(quotes_df, orders_df, qdetect.mapping, odetect.mapping, cfg)
        sheets = {
            "Option A (Rev Match)": result.option_a,
            "Option B (No Rev Match)": result.option_b,
            "Option C (Open Matched)": result.option_c,
            "_Meta": result.meta,
        }
        if cfg.debug and result.debug is not None:
            sheets["_Debug"] = result.debug

        write_output(cfg.out_path, sheets)
        print(f"Wrote output: {cfg.out_path}")
        return 0

    except FollowupError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"Error: unexpected failure ({type(exc).__name__}): {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
