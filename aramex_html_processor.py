"""Prepare Aramex HTML shipment exports for the logistics diagnostic engine.

Aramex portal exports can be raw HTML with several layout/header tables before the
actual shipment rows. This script reads every HTML table with pandas.read_html(),
scores each table for weight-audit columns, cleans the best match, and writes a
standard CSV for the main engine.

Usage:
    python aramex_html_processor.py aramex_export.html cleaned_aramex_shipments.csv
    python aramex_html_processor.py aramex_export.html cleaned.csv --table-index 3
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable

import pandas as pd


CORE_COLUMN_ALIASES = {
    "Waybill_ID": [
        "waybill",
        "waybill no",
        "waybill number",
        "waybill id",
        "awb",
        "awb no",
        "air waybill",
        "shipment number",
        "shipment no",
        "tracking",
        "tracking no",
        "tracking number",
        "parcel number",
        "parcel id",
        "consignment",
        "consignment no",
        "consignment number",
    ],
    "Actual_Weight_KG": [
        "actual weight",
        "actual weight kg",
        "actual wt",
        "actual wt kg",
        "dead weight",
        "dead weight kg",
        "scale weight",
        "physical weight",
        "weight",
        "weight kg",
        "mass",
        "mass kg",
    ],
    "Chargeable_Weight_KG": [
        "chargeable weight",
        "chargeable weight kg",
        "chargeable wt",
        "chargeable wt kg",
        "charged weight",
        "charged weight kg",
        "billing weight",
        "billing weight kg",
        "volumetric weight",
        "volumetric weight kg",
        "vol weight",
        "vol wt",
        "vol kg",
    ],
    "Billed_Weight_KG": [
        "billed weight",
        "billed weight kg",
        "billed wt",
        "billed wt kg",
        "invoice weight",
        "invoice weight kg",
        "invoiced weight",
        "invoiced weight kg",
        "rated weight",
        "rated weight kg",
    ],
}

WEIGHT_COLUMNS = ["Actual_Weight_KG", "Chargeable_Weight_KG", "Billed_Weight_KG"]
MINIMUM_SCORE_TO_AUTO_SELECT = 2


def normalize_label(value: object) -> str:
    """Normalize a messy HTML/header label for matching."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def build_alias_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for standard_name, aliases in CORE_COLUMN_ALIASES.items():
        lookup[normalize_label(standard_name)] = standard_name
        for alias in aliases:
            lookup[normalize_label(alias)] = standard_name
    return lookup


def flatten_columns(columns: Iterable[object]) -> list[str]:
    """Flatten single or multi-level columns into readable strings."""
    flattened = []
    for column in columns:
        if isinstance(column, tuple):
            parts = [str(part).strip() for part in column if str(part).strip() and not str(part).startswith("Unnamed")]
            flattened.append(" ".join(parts))
        else:
            flattened.append(str(column).strip())
    return flattened


def promote_header_row_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """Promote an early row to the header when pandas captured generic numeric headers."""
    cleaned = df.copy()
    cleaned.columns = flatten_columns(cleaned.columns)

    generic_headers = sum(1 for col in cleaned.columns if re.fullmatch(r"Unnamed: \d+|\d+", str(col)))
    if generic_headers < max(1, len(cleaned.columns) // 2):
        return cleaned

    alias_lookup = build_alias_lookup()
    best_row_index = None
    best_row_score = 0

    for row_index in range(min(10, len(cleaned))):
        row_values = [normalize_label(value) for value in cleaned.iloc[row_index].tolist()]
        row_score = sum(1 for value in row_values if value in alias_lookup)
        if row_score > best_row_score:
            best_row_index = row_index
            best_row_score = row_score

    if best_row_index is not None and best_row_score >= 2:
        print(f"Promoting row {best_row_index} to header because it matched {best_row_score} expected columns.")
        cleaned.columns = [str(value).strip() for value in cleaned.iloc[best_row_index].tolist()]
        cleaned = cleaned.iloc[best_row_index + 1 :].reset_index(drop=True)

    return cleaned


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known Aramex/export column variations to engine-friendly names."""
    alias_lookup = build_alias_lookup()
    rename_map = {}
    used_standard_names = set()

    for column in df.columns:
        normalized = normalize_label(column)
        standard_name = alias_lookup.get(normalized)

        # Allow partial matches like "Shipment / Waybill Number".
        if standard_name is None:
            for alias, candidate_standard_name in alias_lookup.items():
                if alias and (alias in normalized or normalized in alias):
                    standard_name = candidate_standard_name
                    break

        if standard_name and standard_name not in used_standard_names:
            rename_map[column] = standard_name
            used_standard_names.add(standard_name)

    return df.rename(columns=rename_map)


def score_table(df: pd.DataFrame) -> tuple[int, list[str]]:
    """Score a candidate table by how many required audit fields it appears to contain."""
    prepared = standardize_columns(promote_header_row_if_needed(df))
    matched_columns = [column for column in CORE_COLUMN_ALIASES if column in prepared.columns]
    row_bonus = 1 if len(prepared.dropna(how="all")) >= 3 else 0
    weight_bonus = 1 if any(column in prepared.columns for column in WEIGHT_COLUMNS) else 0
    score = len(matched_columns) + row_bonus + weight_bonus
    return score, matched_columns


def select_shipment_table(tables: list[pd.DataFrame], table_index: int | None = None) -> pd.DataFrame:
    """Select the most likely shipment table, or use an explicitly supplied index."""
    if table_index is not None:
        try:
            selected = tables[table_index]
        except IndexError as exc:
            raise IndexError(f"Requested table index {table_index}, but only {len(tables)} tables were found.") from exc
        print(f"Using explicit table index {table_index}: shape={selected.shape}")
        return selected

    scored_tables = []
    for index, table in enumerate(tables):
        score, matched_columns = score_table(table)
        scored_tables.append((score, index, matched_columns, table.shape))
        print(f"Table {index}: score={score}, shape={table.shape}, matched={matched_columns}")

    scored_tables.sort(reverse=True, key=lambda item: item[0])
    best_score, best_index, best_matches, best_shape = scored_tables[0]

    if best_score < MINIMUM_SCORE_TO_AUTO_SELECT:
        debug_summary = "\n".join(
            f"  table {index}: score={score}, shape={shape}, matched={matches}"
            for score, index, matches, shape in scored_tables
        )
        raise ValueError(
            "Could not confidently identify the shipment table. "
            "Re-run with --table-index after reviewing this summary:\n" + debug_summary
        )

    print(f"Selected table {best_index}: score={best_score}, shape={best_shape}, matched={best_matches}")
    return tables[best_index]


def clean_weight(value: object) -> float | None:
    """Convert values like '2.5kg', '2,500 g', or '1 234.5 KG' into kilogram floats."""
    if pd.isna(value):
        return None

    text = str(value).strip().lower().replace("\xa0", " ")
    if not text or text in {"nan", "none", "null", "-"}:
        return None

    is_grams = bool(re.search(r"\bgrams?\b|\bg\b", text)) and not re.search(r"\bkg\b|kilograms?", text)
    match = re.search(r"[-+]?\d[\d\s,]*(?:\.\d+)?", text)
    if not match:
        return None

    number_text = match.group(0).replace(" ", "").replace(",", "")
    try:
        value_float = float(number_text)
    except ValueError:
        return None

    return value_float / 1000 if is_grams else value_float


def clean_waybill(value: object) -> str | None:
    """Keep waybill/tracking values stable instead of letting pandas coerce them to floats."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return None
    return re.sub(r"\.0$", "", text)


def clean_selected_table(raw_table: pd.DataFrame) -> pd.DataFrame:
    """Clean headers, isolate core audit columns, and normalize data types."""
    cleaned = promote_header_row_if_needed(raw_table)
    cleaned.columns = flatten_columns(cleaned.columns)
    cleaned = cleaned.dropna(how="all").dropna(axis=1, how="all")
    cleaned = standardize_columns(cleaned)

    available_core_columns = [column for column in CORE_COLUMN_ALIASES if column in cleaned.columns]
    if "Waybill_ID" not in available_core_columns:
        print("WARNING: No waybill/tracking column was matched. Check the source table headers.")
    if not any(column in available_core_columns for column in WEIGHT_COLUMNS):
        print("WARNING: No weight columns were matched. Check aliases or use --table-index if the wrong table was selected.")

    output = cleaned[available_core_columns].copy()

    if "Waybill_ID" in output.columns:
        output["Waybill_ID"] = output["Waybill_ID"].apply(clean_waybill)

    for column in WEIGHT_COLUMNS:
        if column in output.columns:
            output[column] = output[column].apply(clean_weight)

    output = output.dropna(how="all")
    if "Waybill_ID" in output.columns:
        output = output[output["Waybill_ID"].notna()]
        output = output[~output["Waybill_ID"].str.lower().isin(["waybill", "tracking number", "tracking"])]

    return output.reset_index(drop=True)


def process_aramex_html(input_html: str | Path, output_csv: str | Path, table_index: int | None = None) -> pd.DataFrame:
    """Read an Aramex HTML export and write a cleaned CSV for the audit engine."""
    input_path = Path(input_html)
    output_path = Path(output_csv)

    if not input_path.exists():
        raise FileNotFoundError(f"Input HTML file does not exist: {input_path}")

    try:
        tables = pd.read_html(input_path)
    except ValueError as exc:
        raise ValueError(f"No HTML tables found in {input_path}") from exc
    except Exception as exc:
        raise RuntimeError(f"Could not read HTML export {input_path}: {exc}") from exc

    print(f"Found {len(tables)} HTML table(s) in {input_path}")
    selected_table = select_shipment_table(tables, table_index=table_index)
    cleaned = clean_selected_table(selected_table)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    cleaned.to_csv(output_path, index=False)
    print(f"Wrote {len(cleaned)} cleaned shipment row(s) to {output_path}")
    print(f"Final columns: {list(cleaned.columns)}")

    return cleaned


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Clean Aramex HTML shipment exports into CSV for weight audits.")
    parser.add_argument("input_html", help="Path to the local Aramex HTML export file.")
    parser.add_argument("output_csv", help="Path where the cleaned CSV should be written.")
    parser.add_argument(
        "--table-index",
        type=int,
        default=None,
        help="Optional explicit table index if auto-detection picks the wrong table.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        process_aramex_html(args.input_html, args.output_csv, table_index=args.table_index)
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
