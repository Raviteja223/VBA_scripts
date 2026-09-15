from __future__ import annotations

from pathlib import Path
from typing import Any
from difflib import get_close_matches
import unicodedata

import pandas as pd
import yaml


ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"
INPUT_DIR = ROOT / "local_test" / "input"

BUSINESS_SOURCE_COLUMN_KEYS = {
    "status_raw",
    "event_date",
    "settlement_date",
    "return_date",
}

BUSINESS_DERIVED_COLUMN_KEYS = {
    "status",
    "age_days",
    "age_bucket",
    "return_month",
    "return_year",
}

COLUMN_LIST_KEYS = {
    "required_columns",
    "text_columns",
    "business_keys",
    "value_columns",
    "value_cols",
    "date_columns",
}

COLUMN_SINGLE_KEYS = {
    "date_col",
    "date_column",
    "report_date_column",
    "transaction_date_column",
    "status_column",
    "key_column",
    "left_key",
}

IGNORE_CONFIG_KEYS = {
    "output",
    "detail_sheet",
    "remarketing_sheet",
    "era_sheet",
    "required_sheets",
    "source_sheet",
    "filename_regex",
}


def normalize(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = (
        text.replace("\xa0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
    )
    return " ".join(text.split()).strip()


def normalized_key(value: Any) -> str:
    return normalize(value).casefold()


def find_input_workbook() -> Path:
    files = [
        p
        for p in INPUT_DIR.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]

    if not files:
        raise FileNotFoundError(f"No XLSX file found inside:\n{INPUT_DIR}")

    if len(files) > 1:
        print("Multiple XLSX files found:")
        for p in files:
            print(f"  - {p.name}")
        raise RuntimeError(
            "Keep only the workbook being tested inside local_test/input "
            "or temporarily move the other XLSX files out."
        )

    return files[0]


def load_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8-sig") as handle:
        return yaml.safe_load(handle) or {}


def load_all_configs() -> dict[str, dict]:
    result: dict[str, dict] = {}
    for pattern in ("*.yaml", "*.yml"):
        for path in sorted(CONFIG_DIR.glob(pattern)):
            result[path.name] = load_yaml(path)
    return result


def get_source_sheet(configs: dict[str, dict]) -> str:
    workbook_cfg = configs.get("workbook_config.yaml", {})
    return normalize(workbook_cfg.get("source_sheet", "Remarketing"))


def collect_required_sheets(configs: dict[str, dict]) -> list[str]:
    workbook_cfg = configs.get("workbook_config.yaml", {})
    result: list[str] = []

    source = workbook_cfg.get("source_sheet")
    if source:
        result.append(normalize(source))

    for sheet in workbook_cfg.get("required_sheets", []):
        result.append(normalize(sheet))

    return list(dict.fromkeys(result))


def collect_column_refs(
    value: Any,
    path: str = "",
) -> list[tuple[str, str]]:
    """Collect only source/input workbook-column references from a config tree."""
    refs: list[tuple[str, str]] = []

    if isinstance(value, dict):
        for key, child in value.items():
            current_path = f"{path}.{key}" if path else str(key)

            if key in IGNORE_CONFIG_KEYS:
                continue

            if key == "columns" and isinstance(child, dict):
                for logical_name, actual_column in child.items():
                    if logical_name not in BUSINESS_SOURCE_COLUMN_KEYS:
                        continue
                    if isinstance(actual_column, str):
                        refs.append(
                            (f"{current_path}.{logical_name}", actual_column)
                        )
                continue

            if key in COLUMN_LIST_KEYS and isinstance(child, list):
                for index, item in enumerate(child):
                    if isinstance(item, str):
                        refs.append((f"{current_path}[{index}]", item))
                continue

            if key in COLUMN_SINGLE_KEYS and isinstance(child, str):
                refs.append((current_path, child))
                continue

            refs.extend(collect_column_refs(child, current_path))

    elif isinstance(value, list):
        for index, child in enumerate(value):
            refs.extend(collect_column_refs(child, f"{path}[{index}]"))

    return refs


def build_header_lookup(
    columns: list[Any],
) -> tuple[list[str], dict[str, str]]:
    actual_columns = [normalize(c) for c in columns]
    lookup: dict[str, str] = {}
    for column in actual_columns:
        lookup.setdefault(normalized_key(column), column)
    return actual_columns, lookup


def suggest_matches(
    expected: str,
    actual_columns: list[str],
    limit: int = 5,
) -> list[str]:
    normalized_to_actual = {
        normalized_key(c): c for c in actual_columns
    }
    matches = get_close_matches(
        normalized_key(expected),
        list(normalized_to_actual.keys()),
        n=limit,
        cutoff=0.45,
    )
    return [normalized_to_actual[m] for m in matches]


def find_duplicate_normalized_headers(
    columns: list[Any],
) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = {}
    for column in columns:
        raw = str(column)
        groups.setdefault(normalized_key(raw), []).append(raw)
    return {
        key: values
        for key, values in groups.items()
        if len(values) > 1
    }


def derived_column_report(configs: dict[str, dict]) -> list[tuple[str, str]]:
    business_cfg = configs.get("business_rules.yaml", {})
    columns_cfg = business_cfg.get("columns", {})
    result: list[tuple[str, str]] = []
    for logical_name in BUSINESS_DERIVED_COLUMN_KEYS:
        value = columns_cfg.get(logical_name)
        if isinstance(value, str) and value:
            result.append((logical_name, normalize(value)))
    return result


def main() -> None:
    print("=" * 90)
    print("REMARKETING PROJECT - COMPLETE SCHEMA AUDIT")
    print("=" * 90)

    workbook = find_input_workbook()
    configs = load_all_configs()

    print(f"\nWorkbook:\n{workbook}")
    print("\nConfiguration files:")
    for name in configs:
        print(f"  - {name}")

    source_sheet = get_source_sheet(configs)
    print(f"\nConfigured source sheet:\n{source_sheet}")

    excel_file = pd.ExcelFile(workbook, engine="openpyxl")
    workbook_sheets = [normalize(s) for s in excel_file.sheet_names]
    print(f"\nWorkbook sheet count: {len(workbook_sheets)}")

    print("\n" + "=" * 90)
    print("1. SHEET AUDIT")
    print("=" * 90)

    required_sheets = collect_required_sheets(configs)
    normalized_sheet_lookup = {
        normalized_key(s): s for s in workbook_sheets
    }
    missing_sheets: list[str] = []

    for sheet in required_sheets:
        real = normalized_sheet_lookup.get(normalized_key(sheet))
        if real is None:
            print(f"MISSING    {sheet}")
            missing_sheets.append(sheet)
        else:
            print(f"OK         {real}")

    real_source_sheet = normalized_sheet_lookup.get(
        normalized_key(source_sheet)
    )
    if real_source_sheet is None:
        raise KeyError(
            f"Configured source sheet {source_sheet!r} was not found."
        )
    source_sheet = real_source_sheet

    df = pd.read_excel(
        workbook,
        sheet_name=source_sheet,
        nrows=0,
        engine="openpyxl",
    )

    actual_columns, header_lookup = build_header_lookup(list(df.columns))

    print("\n" + "=" * 90)
    print("2. SOURCE COLUMN INVENTORY")
    print("=" * 90)
    print(f"\n{source_sheet}: {len(actual_columns)} columns")
    for index, column in enumerate(actual_columns):
        print(f"{index:>3}: {column!r}")

    duplicates = find_duplicate_normalized_headers(list(df.columns))

    print("\n" + "=" * 90)
    print("3. DUPLICATE / AMBIGUOUS HEADERS")
    print("=" * 90)
    if not duplicates:
        print("None")
    else:
        for key, values in duplicates.items():
            print(f"{key!r}: {values}")

    references: list[tuple[str, str, str]] = []
    for config_name, config in configs.items():
        for config_path, column in collect_column_refs(config):
            references.append(
                (config_name, config_path, normalize(column))
            )

    # de-duplicate while preserving order
    seen: set[tuple[str, str, str]] = set()
    unique_refs: list[tuple[str, str, str]] = []
    for item in references:
        if item not in seen:
            seen.add(item)
            unique_refs.append(item)

    print("\n" + "=" * 90)
    print("4. CONFIGURATION SOURCE-COLUMN AUDIT")
    print("=" * 90)

    exact_matches: list[tuple[str, str, str]] = []
    normalized_matches: list[tuple[str, str, str, str]] = []
    errors: list[tuple[str, str, str, list[str]]] = []
    actual_exact = set(actual_columns)

    for config_name, config_path, expected in unique_refs:
        if expected in actual_exact:
            exact_matches.append((config_name, config_path, expected))
            print(
                f"OK          {config_name} :: {config_path} "
                f"-> {expected!r}"
            )
            continue

        normalized = header_lookup.get(normalized_key(expected))
        if normalized is not None:
            normalized_matches.append(
                (config_name, config_path, expected, normalized)
            )
            print(
                f"NORMALIZED  {config_name} :: {config_path} "
                f"-> {expected!r} => {normalized!r}"
            )
            continue

        suggestions = suggest_matches(expected, actual_columns)
        errors.append(
            (config_name, config_path, expected, suggestions)
        )
        print(
            f"MISSING     {config_name} :: {config_path} "
            f"-> {expected!r}"
        )
        if suggestions:
            print("            Suggestions only:")
            for suggestion in suggestions:
                print(f"              - {suggestion!r}")

    print("\n" + "=" * 90)
    print("5. DERIVED COLUMNS (NOT REQUIRED IN SOURCE XLSX)")
    print("=" * 90)
    derived = derived_column_report(configs)
    if not derived:
        print("None configured")
    else:
        for logical_name, actual_name in derived:
            print(f"GENERATED   {logical_name:<15} -> {actual_name!r}")

    print("\n" + "=" * 90)
    print("6. AUDIT SUMMARY")
    print("=" * 90)
    print(
        f"""
Required sheets checked : {len(required_sheets)}
Missing sheets          : {len(missing_sheets)}

Configured source refs  : {len(unique_refs)}
Exact matches           : {len(exact_matches)}
Normalized matches      : {len(normalized_matches)}
Missing source refs     : {len(errors)}

Derived output columns  : {len(derived)}
Duplicate headers       : {len(duplicates)}
"""
    )

    if errors:
        print("=" * 90)
        print("7. CONFIG VALUES REQUIRING REVIEW")
        print("=" * 90)
        for config_name, config_path, expected, suggestions in errors:
            print(f"\nFile       : {config_name}")
            print(f"Config path: {config_path}")
            print(f"Configured : {expected!r}")
            if suggestions:
                print("Possible workbook columns (suggestions only):")
                for suggestion in suggestions:
                    print(f"  - {suggestion!r}")
            else:
                print("No similar workbook column found.")

        print(
            "\nDo not run the main processing pipeline until the "
            "missing SOURCE references above are reviewed."
        )
        raise SystemExit(1)

    if missing_sheets:
        raise SystemExit(1)

    if duplicates:
        raise SystemExit(
            "Resolve duplicate normalized headers before processing."
        )

    print("\nSCHEMA AUDIT PASSED.")
    print(
        "All configured SOURCE columns are compatible with the "
        "source workbook. Derived columns will be validated after "
        "the calculation step."
    )


if __name__ == "__main__":
    main()
