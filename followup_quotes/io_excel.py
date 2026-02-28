from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable

from openpyxl import Workbook, load_workbook
import pandas as pd

from .config import FollowupError


PUNCT_RE = re.compile(r"[\W_]+", flags=re.UNICODE)


@dataclass
class DetectionResult:
    mapping: dict[str, str]
    notes: list[str]


def normalize_header(value: str) -> str:
    return " ".join(str(value).strip().lower().split())


def normalize_customer(value: object) -> str:
    if pd.isna(value):
        return ""
    s = str(value).upper()
    s = PUNCT_RE.sub("", s)
    return s


def parse_truthy(value: object) -> bool:
    if pd.isna(value):
        return False
    token = str(value).strip().upper()
    return token in {"TRUE", "1", "YES", "Y", "T"}


def parse_money(value: object) -> float | None:
    if pd.isna(value):
        return None
    s = str(value).strip().replace(",", "").replace("$", "")
    try:
        return float(s)
    except ValueError:
        return None


def read_excel(path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name or 0, dtype=object)


def _find_header(headers: list[str], synonyms: Iterable[str], contains: str | None = None) -> str | None:
    normalized = {normalize_header(h): h for h in headers}
    for cand in synonyms:
        n = normalize_header(cand)
        if n in normalized:
            return normalized[n]
    if contains:
        for h in headers:
            if contains in normalize_header(h):
                return h
    return None


def detect_columns(
    df: pd.DataFrame,
    synonyms: dict[str, list[str]],
    required_fields: set[str],
    overrides: dict[str, str] | None = None,
    contains_rules: dict[str, str] | None = None,
) -> DetectionResult:
    overrides = overrides or {}
    contains_rules = contains_rules or {}
    headers = [str(c) for c in df.columns]
    mapping: dict[str, str] = {}
    notes: list[str] = []

    for field, col in overrides.items():
        if col in headers:
            mapping[field] = col
            notes.append(f"override: {field} -> {col}")

    for field, syns in synonyms.items():
        if field in mapping:
            continue
        found = _find_header(headers, syns, contains_rules.get(field))
        if found:
            mapping[field] = found

    missing = [f for f in required_fields if f not in mapping]
    if missing:
        detail = []
        for field in missing:
            tried = synonyms.get(field, [])
            detail.append(f"- {field}: tried {tried}")
        raise FollowupError(
            "Required field detection failed.\n"
            + "Missing fields:\n"
            + "\n".join(f"- {m}" for m in missing)
            + "\nExisting columns:\n"
            + "\n".join(f"- {h}" for h in headers)
            + "\nSynonyms tried:\n"
            + "\n".join(detail)
            + "\nUse --column-map mapping.json with {'quotes': {...}, 'orders': {...}} to override."
        )

    return DetectionResult(mapping=mapping, notes=notes)


def safe_excel_value(value: object) -> object:
    if isinstance(value, str) and value[:1] in {"=", "+", "-", "@"}:
        return "'" + value
    return value


def _sheet_header_row(sheet, columns: list[str], scan_rows: int = 60) -> int:
    target = [normalize_header(c) for c in columns]
    for r in range(1, min(scan_rows, sheet.max_row) + 1):
        values = [normalize_header(sheet.cell(row=r, column=c).value) for c in range(1, len(columns) + 1)]
        if values == target:
            return r
    return 1


def _write_dataframe_to_sheet(sheet, df: pd.DataFrame) -> None:
    cols = [str(c) for c in df.columns]
    header_row = _sheet_header_row(sheet, cols)

    max_cols = max(len(cols), sheet.max_column)
    for r in range(header_row + 1, sheet.max_row + 1):
        for c in range(1, max_cols + 1):
            sheet.cell(row=r, column=c).value = None

    for c, col in enumerate(cols, start=1):
        sheet.cell(row=header_row, column=c).value = col

    for ridx, row in enumerate(df.itertuples(index=False, name=None), start=header_row + 1):
        for cidx, value in enumerate(row, start=1):
            sheet.cell(row=ridx, column=cidx).value = safe_excel_value(value)


def write_output(path: Path, sheets: dict[str, pd.DataFrame], template_path: Path | None = None) -> None:
    if template_path:
        wb = load_workbook(template_path)
    else:
        wb = Workbook()
        default = wb.active
        wb.remove(default)

    for sheet_name, df in sheets.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
        _write_dataframe_to_sheet(ws, df)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
