from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .app import generate_followup_workbook
from .config import ColumnMap, FollowupError, load_reps, RunConfig


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

        out = generate_followup_workbook(cfg)
        print(f"Wrote output: {out}")
        return 0

    except FollowupError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"Error: unexpected failure ({type(exc).__name__}): {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
