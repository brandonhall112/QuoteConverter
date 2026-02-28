from __future__ import annotations

from pathlib import Path
import re

from .config import ColumnMap, DEFAULT_ALLOWED_REPS, ORDER_SYNONYMS, QUOTE_SYNONYMS, RunConfig
from .io_excel import detect_columns, read_excel, write_output
from .matching import run_matching

INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


def _sheet_name_for_rep(rep: str) -> str:
    clean = INVALID_SHEET_CHARS.sub("-", rep).strip() or "Unassigned"
    return clean[:31]


def generate_followup_workbook(cfg: RunConfig) -> Path:
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

    result = run_matching(quotes_df, orders_df, qdetect.mapping, odetect.mapping, cfg)
    sheets = {
        "Follow-Up": result.followups,
        "_Meta": result.meta,
    }

    for rep, rep_df in result.followups.groupby("Entry Person Name", dropna=False):
        rep_name = "Unassigned" if rep is None or str(rep).strip() == "" else str(rep)
        sheets[_sheet_name_for_rep(rep_name)] = rep_df.reset_index(drop=True)

    if cfg.debug and result.debug is not None:
        sheets["_Debug"] = result.debug

    write_output(cfg.out_path, sheets, cfg.template_path)
    return cfg.out_path


def make_run_config(
    quotes: str,
    orders: str,
    out: str,
    *,
    floor: float = 1500,
    tolerance: float = 1,
    relative_tolerance: float = 0.05,
    sheet_quotes: str | None = None,
    sheet_orders: str | None = None,
    reps: list[str] | None = None,
    debug: bool = False,
    fuzzy: bool = False,
    fuzzy_threshold: int = 90,
    column_map: ColumnMap | None = None,
    template: str | None = None,
) -> RunConfig:
    return RunConfig(
        quotes_path=Path(quotes),
        orders_path=Path(orders),
        out_path=Path(out),
        floor=floor,
        tolerance=tolerance,
        relative_tolerance=relative_tolerance,
        sheet_quotes=sheet_quotes,
        sheet_orders=sheet_orders,
        reps=reps if reps is not None else DEFAULT_ALLOWED_REPS.copy(),
        debug=debug,
        fuzzy=fuzzy,
        fuzzy_threshold=fuzzy_threshold,
        column_map=column_map or ColumnMap(),
        template_path=Path(template) if template else None,
    )
