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


# ---------------------------------------------------------------------
# NORMALIZATION
# ---------------------------------------------------------------------

def normalize(value: Any) -> str:
    if value is None:
        return ""

    text = unicodedata.normalize("NFKC", str(value))

    text = (
        text
        .replace("\xa0", " ")
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


# ---------------------------------------------------------------------
# FILE DISCOVERY
# ---------------------------------------------------------------------

def find_input_workbook() -> Path:
    files = [
        p
        for p in INPUT_DIR.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]

    if not files:
        raise FileNotFoundError(
            f"No XLSX file found inside:\n{INPUT_DIR}"
        )

    if len(files) > 1:
        print("Multiple XLSX files found:")
        for p in files:
            print(f"  - {p.name}")

        raise RuntimeError(
            "Keep only the workbook being tested inside local_test/input."
        )

    return files[0]


# ---------------------------------------------------------------------
# YAML
# ---------------------------------------------------------------------

def load_yaml(path: Path) -> dict:
    with path.open(
        "r",
        encoding="utf-8-sig",
    ) as f:
        data = yaml.safe_load(f)

    return data or {}


def load_all_configs() -> dict[str, dict]:
    result = {}

    for path in sorted(CONFIG_DIR.glob("*.yaml")):
        result[path.name] = load_yaml(path)

    for path in sorted(CONFIG_DIR.glob("*.yml")):
        result[path.name] = load_yaml(path)

    return result


# ---------------------------------------------------------------------
# DETERMINE SOURCE SHEET
# ---------------------------------------------------------------------

def get_source_sheet(configs: dict[str, dict]) -> str:
    workbook_cfg = configs.get(
        "workbook_config.yaml",
        {},
    )

    source_sheet = workbook_cfg.get(
        "source_sheet",
        "Remarketing",
    )

    return normalize(source_sheet)


# ---------------------------------------------------------------------
# EXTRACT CONFIGURED SOURCE COLUMN REFERENCES
# ---------------------------------------------------------------------

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


def collect_column_refs(
    value: Any,
    path: str = "",
    parent_key: str | None = None,
) -> list[tuple[str, str]]:
    """
    Returns:
        [
          ("business_rules.yaml.columns.event_date", "Date de retour"),
          ...
        ]
    """

    refs: list[tuple[str, str]] = []

    if isinstance(value, dict):

        for key, child in value.items():

            current_path = (
                f"{path}.{key}"
                if path
                else str(key)
            )

            if key in IGNORE_CONFIG_KEYS:
                continue

            # business_rules.yaml -> columns -> logical_name: actual header
            if key == "columns" and isinstance(child, dict):

                for logical_name, actual_column in child.items():

                    if isinstance(actual_column, str):
                        refs.append(
                            (
                                f"{current_path}.{logical_name}",
                                actual_column,
                            )
                        )

                continue

            # Lists known to contain workbook columns.
            if key in COLUMN_LIST_KEYS and isinstance(child, list):

                for index, item in enumerate(child):

                    if isinstance(item, str):
                        refs.append(
                            (
                                f"{current_path}[{index}]",
                                item,
                            )
                        )

                continue

            # Single configured column.
            if key in COLUMN_SINGLE_KEYS and isinstance(child, str):

                refs.append(
                    (
                        current_path,
                        child,
                    )
                )

                continue

            refs.extend(
                collect_column_refs(
                    child,
                    path=current_path,
                    parent_key=key,
                )
            )

    elif isinstance(value, list):

        for index, child in enumerate(value):

            refs.extend(
                collect_column_refs(
                    child,
                    path=f"{path}[{index}]",
                    parent_key=parent_key,
                )
            )

    return refs


# ---------------------------------------------------------------------
# SHEET REFERENCES
# ---------------------------------------------------------------------

def collect_required_sheets(
    configs: dict[str, dict],
) -> list[str]:

    workbook_cfg = configs.get(
        "workbook_config.yaml",
        {},
    )

    result = []

    source = workbook_cfg.get("source_sheet")

    if source:
        result.append(normalize(source))

    for sheet in workbook_cfg.get(
        "required_sheets",
        [],
    ):
        result.append(normalize(sheet))

    return list(dict.fromkeys(result))


# ---------------------------------------------------------------------
# MATCHING
# ---------------------------------------------------------------------

def build_header_lookup(
    columns: list[Any],
) -> tuple[
    list[str],
    dict[str, str],
]:
    actual_columns = [
        normalize(c)
        for c in columns
    ]

    lookup = {}

    for column in actual_columns:

        key = normalized_key(column)

        if key not in lookup:
            lookup[key] = column

    return actual_columns, lookup


def suggest_matches(
    expected: str,
    actual_columns: list[str],
    limit: int = 5,
) -> list[str]:

    normalized_to_actual = {
        normalized_key(c): c
        for c in actual_columns
    }

    matches = get_close_matches(
        normalized_key(expected),
        list(normalized_to_actual.keys()),
        n=limit,
        cutoff=0.45,
    )

    return [
        normalized_to_actual[m]
        for m in matches
    ]


# ---------------------------------------------------------------------
# DUPLICATE HEADER CHECK
# ---------------------------------------------------------------------

def find_duplicate_normalized_headers(
    columns: list[Any],
) -> dict[str, list[str]]:

    groups: dict[str, list[str]] = {}

    for column in columns:

        raw = str(column)
        key = normalized_key(raw)

        groups.setdefault(
            key,
            [],
        ).append(raw)

    return {
        key: values
        for key, values in groups.items()
        if len(values) > 1
    }


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

def main():

    print("=" * 90)
    print("REMARKETING PROJECT - COMPLETE SCHEMA AUDIT")
    print("=" * 90)

    workbook = find_input_workbook()

    print(f"\nWorkbook:")
    print(workbook)

    configs = load_all_configs()

    print("\nConfiguration files:")

    for name in configs:
        print(f"  - {name}")

    source_sheet = get_source_sheet(configs)

    print(f"\nConfigured source sheet:")
    print(source_sheet)

    # ---------------------------------------------------------
    # READ WORKBOOK METADATA
    # ---------------------------------------------------------

    excel_file = pd.ExcelFile(
        workbook,
        engine="openpyxl",
    )

    workbook_sheets = [
        normalize(s)
        for s in excel_file.sheet_names
    ]

    print(f"\nWorkbook sheet count: {len(workbook_sheets)}")

    # ---------------------------------------------------------
    # SHEET AUDIT
    # ---------------------------------------------------------

    print("\n" + "=" * 90)
    print("1. SHEET AUDIT")
    print("=" * 90)

    required_sheets = collect_required_sheets(
        configs
    )

    sheet_errors = []

    for sheet in required_sheets:

        exists = (
            normalized_key(sheet)
            in {
                normalized_key(x)
                for x in workbook_sheets
            }
        )

        status = "OK" if exists else "MISSING"

        print(
            f"{status:<10} {sheet}"
        )

        if not exists:
            sheet_errors.append(sheet)

    if source_sheet not in workbook_sheets:

        normalized_sheets = {
            normalized_key(s): s
            for s in workbook_sheets
        }

        real_source = normalized_sheets.get(
            normalized_key(source_sheet)
        )

        if real_source:
            source_sheet = real_source
        else:
            raise KeyError(
                f"Source sheet {source_sheet!r} "
                f"was not found."
            )

    # ---------------------------------------------------------
    # LOAD ONLY HEADER
    # ---------------------------------------------------------

    df = pd.read_excel(
        workbook,
        sheet_name=source_sheet,
        nrows=0,
        engine="openpyxl",
    )

    actual_columns, header_lookup = (
        build_header_lookup(
            list(df.columns)
        )
    )

    print("\n" + "=" * 90)
    print("2. SOURCE COLUMN INVENTORY")
    print("=" * 90)

    print(
        f"\n{source_sheet}: "
        f"{len(actual_columns)} columns"
    )

    for index, column in enumerate(
        actual_columns
    ):
        print(
            f"{index:>3}: {column!r}"
        )

    # ---------------------------------------------------------
    # DUPLICATE HEADERS
    # ---------------------------------------------------------

    duplicates = (
        find_duplicate_normalized_headers(
            list(df.columns)
        )
    )

    print("\n" + "=" * 90)
    print("3. DUPLICATE / AMBIGUOUS HEADERS")
    print("=" * 90)

    if not duplicates:

        print("None")

    else:

        for key, values in duplicates.items():

            print(
                f"{key!r}: {values}"
            )

    # ---------------------------------------------------------
    # COLLECT EVERY CONFIG COLUMN
    # ---------------------------------------------------------

    references = []

    for config_name, config in configs.items():

        refs = collect_column_refs(
            config
        )

        for path, column in refs:

            references.append(
                (
                    config_name,
                    path,
                    normalize(column),
                )
            )

    # Remove duplicates but keep paths.
    seen = set()
    unique_refs = []

    for item in references:

        if item not in seen:
            seen.add(item)
            unique_refs.append(item)

    # ---------------------------------------------------------
    # COLUMN AUDIT
    # ---------------------------------------------------------

    print("\n" + "=" * 90)
    print("4. CONFIGURATION COLUMN AUDIT")
    print("=" * 90)

    errors = []
    normalized_matches = []
    exact_matches = []

    actual_exact = set(actual_columns)

    for (
        config_name,
        config_path,
        expected,
    ) in unique_refs:

        if expected in actual_exact:

            exact_matches.append(
                (
                    config_name,
                    config_path,
                    expected,
                )
            )

            print(
                f"OK          "
                f"{config_name} :: "
                f"{config_path} "
                f"-> {expected!r}"
            )

            continue

        normalized = header_lookup.get(
            normalized_key(expected)
        )

        if normalized is not None:

            normalized_matches.append(
                (
                    config_name,
                    config_path,
                    expected,
                    normalized,
                )
            )

            print(
                f"NORMALIZED  "
                f"{config_name} :: "
                f"{config_path} "
                f"-> {expected!r} "
                f"=> {normalized!r}"
            )

            continue

        suggestions = suggest_matches(
            expected,
            actual_columns,
        )

        errors.append(
            (
                config_name,
                config_path,
                expected,
                suggestions,
            )
        )

        print(
            f"MISSING     "
            f"{config_name} :: "
            f"{config_path} "
            f"-> {expected!r}"
        )

        if suggestions:

            print(
                "            Suggestions only:"
            )

            for suggestion in suggestions:
                print(
                    f"              - {suggestion!r}"
                )

    # ---------------------------------------------------------
    # SUMMARY
    # ---------------------------------------------------------

    print("\n" + "=" * 90)
    print("5. AUDIT SUMMARY")
    print("=" * 90)

    print(
        f"""
Required sheets checked : {len(required_sheets)}
Missing sheets          : {len(sheet_errors)}

Configured column refs  : {len(unique_refs)}
Exact matches           : {len(exact_matches)}
Normalized matches      : {len(normalized_matches)}
Missing references      : {len(errors)}

Duplicate headers       : {len(duplicates)}
"""
    )

    # ---------------------------------------------------------
    # FIX LIST
    # ---------------------------------------------------------

    if errors:

        print("=" * 90)
        print("6. CONFIG VALUES REQUIRING REVIEW")
        print("=" * 90)

        for (
            config_name,
            config_path,
            expected,
            suggestions,
        ) in errors:

            print(
                f"\nFile       : {config_name}"
            )

            print(
                f"Config path: {config_path}"
            )

            print(
                f"Configured : {expected!r}"
            )

            if suggestions:

                print(
                    "Possible workbook columns:"
                )

                for suggestion in suggestions:

                    print(
                        f"  - {suggestion!r}"
                    )

            else:

                print(
                    "No similar workbook column found."
                )

        print(
            "\nDO NOT run the main processing pipeline "
            "until these references have been reviewed."
        )

        raise SystemExit(1)

    if sheet_errors:

        raise SystemExit(1)

    if duplicates:

        raise SystemExit(
            "Resolve duplicate normalized headers first."
        )

    print(
        "\nSCHEMA AUDIT PASSED."
    )

    print(
        "Configuration column references are compatible "
        "with the source workbook."
    )


if __name__ == "__main__":
    main()
