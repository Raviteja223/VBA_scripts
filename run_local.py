from __future__ import annotations

import copy
import os
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from include.remarketing.excel_reader import (
    read_configured,
    read_sheet,
    sheet_names,
)
from include.remarketing.workbook_validation import validate_workbook
from include.remarketing.standardization import standardize
from include.remarketing.mappings import left_lookup
from include.remarketing.calculations import apply_rules
from include.remarketing.monthly_reporting import monthly_metrics
from include.remarketing.ytd_reporting import ytd_slice
from include.remarketing.report_builder import build_report
from include.remarketing.reconciliation import reconcile


ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"
INPUT_DIR = ROOT / "local_test" / "input"
OUTPUT_DIR = ROOT / "local_test" / "output"
TEMP_DIR = ROOT / "local_test" / "temp"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

BUSINESS_SOURCE_COLUMN_KEYS = (
    "status_raw",
    "event_date",
    "settlement_date",
    "return_date",
)

BUSINESS_DERIVED_COLUMN_KEYS = (
    "status",
    "age_days",
    "age_bucket",
    "return_month",
    "return_year",
)


# -----------------------------------------------------------------------------
# NORMALIZATION / RESOLUTION
# -----------------------------------------------------------------------------


def normalize_name(value: Any) -> str:
    """Return a display-safe normalized name without changing business meaning."""
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


def normalize_key(value: Any) -> str:
    """Return a comparison-only normalized key."""
    return normalize_name(value).casefold()


def build_name_lookup(values: list[Any]) -> dict[str, str]:
    """Map normalized name -> actual name, while rejecting ambiguity."""
    lookup: dict[str, str] = {}

    for value in values:
        actual = str(value)
        key = normalize_key(actual)

        if key in lookup and lookup[key] != actual:
            raise ValueError(
                "Ambiguous names after normalization: "
                f"{lookup[key]!r} and {actual!r}"
            )

        lookup[key] = actual

    return lookup


def build_column_lookup(df: pd.DataFrame) -> dict[str, str]:
    return build_name_lookup([str(column) for column in df.columns])


def resolve_actual_column(
    df: pd.DataFrame,
    configured_name: str,
    *,
    label: str = "column",
) -> str:
    lookup = build_column_lookup(df)
    actual = lookup.get(normalize_key(configured_name))

    if actual is None:
        raise KeyError(
            f"Configured {label} {configured_name!r} does not exist in "
            "the processing dataframe."
        )

    return actual


def resolve_optional_column(
    df: pd.DataFrame,
    configured_name: str | None,
    *,
    label: str = "column",
) -> str | None:
    if not configured_name:
        return None
    return resolve_actual_column(df, configured_name, label=label)


def resolve_column_list(
    df: pd.DataFrame,
    configured_columns: list[str] | str | None,
    *,
    label: str,
) -> list[str]:
    if not configured_columns:
        return []

    if isinstance(configured_columns, str):
        configured_columns = [configured_columns]

    resolved: list[str] = []
    seen: set[str] = set()

    for configured in configured_columns:
        actual = resolve_actual_column(df, str(configured), label=label)
        key = normalize_key(actual)
        if key not in seen:
            seen.add(key)
            resolved.append(actual)

    return resolved


def resolve_sheet_name(
    input_file: Path,
    configured_name: str,
) -> str:
    available = sheet_names(input_file)
    lookup = build_name_lookup(available)
    actual = lookup.get(normalize_key(configured_name))

    if actual is None:
        raise KeyError(
            f"Configured worksheet {configured_name!r} was not found. "
            f"Available sheets: {available}"
        )

    return actual


# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}

    with path.open("r", encoding="utf-8-sig") as handle:
        return yaml.safe_load(handle) or {}


def load_configs() -> dict[str, dict[str, Any]]:
    return {
        "workbook": load_yaml(CONFIG_DIR / "workbook_config.yaml"),
        "business": load_yaml(CONFIG_DIR / "business_rules.yaml"),
        "columns": load_yaml(CONFIG_DIR / "column_mapping.yaml"),
        "report": load_yaml(CONFIG_DIR / "report_config.yaml"),
    }


# -----------------------------------------------------------------------------
# INPUT FILE
# -----------------------------------------------------------------------------


def resolve_input_file() -> Path:
    explicit = os.getenv("LOCAL_INPUT_FILE")

    if explicit:
        path = Path(explicit).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(
                f"LOCAL_INPUT_FILE does not exist:\n{path}"
            )
        return path

    candidates = [
        path
        for path in INPUT_DIR.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]

    if not candidates:
        raise FileNotFoundError(
            f"No XLSX file found inside:\n{INPUT_DIR}"
        )

    if len(candidates) > 1:
        names = "\n".join(f"  - {p.name}" for p in candidates)
        raise RuntimeError(
            "More than one XLSX file exists in local_test/input.\n"
            "Either keep only the workbook being tested or set "
            "LOCAL_INPUT_FILE.\n\n"
            f"Files found:\n{names}"
        )

    return candidates[0]


# -----------------------------------------------------------------------------
# SOURCE / DERIVED COLUMN VALIDATION
# -----------------------------------------------------------------------------


def configured_source_columns(
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> list[str]:
    required: list[str] = []

    for key in ("required_columns", "business_keys", "text_columns"):
        values = workbook_cfg.get(key, []) or []
        if isinstance(values, str):
            values = [values]
        required.extend(str(value) for value in values if value)

    columns_cfg = business_cfg.get("columns", {}) or {}
    for logical_key in BUSINESS_SOURCE_COLUMN_KEYS:
        value = columns_cfg.get(logical_key)
        if value:
            required.append(str(value))

    unique: list[str] = []
    seen: set[str] = set()

    for configured in required:
        key = normalize_key(configured)
        if key not in seen:
            seen.add(key)
            unique.append(configured)

    return unique


def validate_required_source_columns(
    df: pd.DataFrame,
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> dict[str, str]:
    """Validate source-only columns and return config-name -> actual-name."""
    lookup = build_column_lookup(df)
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for configured in configured_source_columns(workbook_cfg, business_cfg):
        actual = lookup.get(normalize_key(configured))
        if actual is None:
            missing.append(configured)
        else:
            resolved[configured] = actual

    if missing:
        available = "\n".join(f"  - {str(column)!r}" for column in df.columns)
        raise KeyError(
            "Required SOURCE column(s) missing from the processing sheet: "
            + ", ".join(missing)
            + "\n\nAvailable source columns:\n"
            + available
        )

    print("Required source-column validation passed.")
    print("Resolved source columns:")
    for configured, actual in resolved.items():
        if configured == actual:
            print(f"  - {configured!r}")
        else:
            print(f"  - {configured!r} -> {actual!r}")

    return resolved


def validate_derived_columns(
    df: pd.DataFrame,
    business_cfg: dict[str, Any],
) -> dict[str, str]:
    """Validate fields that apply_rules() is expected to generate."""
    columns_cfg = business_cfg.get("columns", {}) or {}
    expected = [
        str(columns_cfg[key])
        for key in BUSINESS_DERIVED_COLUMN_KEYS
        if columns_cfg.get(key)
    ]

    lookup = build_column_lookup(df)
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for configured in expected:
        actual = lookup.get(normalize_key(configured))
        if actual is None:
            missing.append(configured)
        else:
            resolved[configured] = actual

    if missing:
        raise KeyError(
            "Calculation step did not create expected derived column(s): "
            + ", ".join(missing)
        )

    print("Derived-column validation passed.")
    print("Generated columns:")
    for configured, actual in resolved.items():
        if configured == actual:
            print(f"  - {configured!r}")
        else:
            print(f"  - {configured!r} -> {actual!r}")

    return resolved


def resolve_runtime_business_config(
    df: pd.DataFrame,
    business_cfg: dict[str, Any],
) -> dict[str, Any]:
    """
    Replace configured SOURCE-column names with actual dataframe column names.

    Derived output names are intentionally kept exactly as configured because
    calculations.apply_rules() is expected to create them.
    """
    runtime_cfg = copy.deepcopy(business_cfg)
    columns_cfg = runtime_cfg.setdefault("columns", {})

    for logical_key in BUSINESS_SOURCE_COLUMN_KEYS:
        configured = columns_cfg.get(logical_key)
        if configured:
            columns_cfg[logical_key] = resolve_actual_column(
                df,
                str(configured),
                label=f"business_rules.columns.{logical_key}",
            )

    return runtime_cfg


# -----------------------------------------------------------------------------
# PERIOD
# -----------------------------------------------------------------------------


def resolve_period(
    df: pd.DataFrame,
    business_cfg: dict[str, Any],
) -> str:
    """Resolve YYYY-MM for local execution."""
    explicit = os.getenv("TEST_PERIOD")

    if explicit:
        try:
            period = pd.Period(explicit, freq="M")
        except Exception as exc:
            raise ValueError(
                "TEST_PERIOD must be in YYYY-MM format, e.g. 2026-07"
            ) from exc
        return str(period)

    columns_cfg = business_cfg.get("columns", {}) or {}
    event_date = columns_cfg.get("event_date")

    if not event_date:
        raise KeyError(
            "business_rules.yaml is missing columns.event_date. "
            "Set TEST_PERIOD or configure columns.event_date."
        )

    event_date = resolve_actual_column(
        df,
        str(event_date),
        label="event_date column",
    )

    dates = pd.to_datetime(
        df[event_date],
        errors="coerce",
        dayfirst=True,
    ).dropna()

    if dates.empty:
        raise ValueError(
            f"Could not infer a reporting period because {event_date!r} "
            "contains no parseable dates. Set TEST_PERIOD explicitly."
        )

    return str(dates.max().to_period("M"))


# -----------------------------------------------------------------------------
# REFERENCE LOOKUPS
# -----------------------------------------------------------------------------


def _lookup_definitions(mapping_cfg: dict[str, Any]) -> list[dict[str, Any]]:
    if not mapping_cfg:
        return []

    for key in ("lookups", "mappings", "reference_lookups"):
        value = mapping_cfg.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]

    return []


def apply_reference_lookups(
    df: pd.DataFrame,
    input_file: Path,
    mapping_cfg: dict[str, Any],
) -> pd.DataFrame:
    """Apply explicitly configured workbook lookups with normalized resolution."""
    result = df
    definitions = _lookup_definitions(mapping_cfg)

    if not definitions:
        print("No configured reference lookups found; skipping lookup stage.")
        return result

    available_sheets = sheet_names(input_file)
    sheet_lookup = build_name_lookup(available_sheets)

    for index, definition in enumerate(definitions, start=1):
        configured_sheet = definition.get("sheet") or definition.get("reference_sheet")
        configured_left_key = definition.get("left_key")
        configured_right_key = definition.get("right_key")
        value_columns = definition.get("value_columns") or definition.get("columns") or []
        required = bool(definition.get("required", False))

        if isinstance(value_columns, str):
            value_columns = [value_columns]

        missing_config = [
            name
            for name, value in {
                "sheet": configured_sheet,
                "left_key": configured_left_key,
                "right_key": configured_right_key,
            }.items()
            if not value
        ]

        if missing_config:
            raise ValueError(
                f"Lookup #{index} is missing config field(s): "
                + ", ".join(missing_config)
            )

        actual_sheet = sheet_lookup.get(normalize_key(configured_sheet))
        if actual_sheet is None:
            raise KeyError(
                f"Lookup #{index} reference sheet {configured_sheet!r} does not exist."
            )

        actual_left_key = resolve_actual_column(
            result,
            str(configured_left_key),
            label=f"lookup #{index} left_key",
        )

        reference = read_sheet(input_file, actual_sheet)
        actual_right_key = resolve_actual_column(
            reference,
            str(configured_right_key),
            label=f"lookup #{index} right_key",
        )
        actual_value_columns = resolve_column_list(
            reference,
            [str(column) for column in value_columns],
            label=f"lookup #{index} value column",
        )

        result = left_lookup(
            result,
            reference,
            left_key=actual_left_key,
            right_key=actual_right_key,
            value_columns=actual_value_columns,
            required=required,
        )

        print(
            f"Applied lookup #{index}: {actual_sheet!r} "
            f"({actual_left_key!r} -> {actual_right_key!r})"
        )

    return result


# -----------------------------------------------------------------------------
# REPORT SETTINGS
# -----------------------------------------------------------------------------


def find_list_by_keys(
    config: dict[str, Any],
    keys: tuple[str, ...],
) -> list[str]:
    """Recursively find the first list-of-strings under any candidate key."""

    def walk(value: Any) -> list[str] | None:
        if isinstance(value, dict):
            for key, child in value.items():
                if key in keys:
                    if isinstance(child, str):
                        return [child]
                    if isinstance(child, list):
                        return [str(x) for x in child if isinstance(x, str)]

            for child in value.values():
                found = walk(child)
                if found is not None:
                    return found

        elif isinstance(value, list):
            for child in value:
                found = walk(child)
                if found is not None:
                    return found

        return None

    return walk(config) or []


def resolve_value_columns(
    df: pd.DataFrame,
    report_cfg: dict[str, Any],
) -> list[str]:
    configured = find_list_by_keys(
        report_cfg,
        (
            "value_cols",
            "value_columns",
            "monthly_value_cols",
            "monthly_value_columns",
        ),
    )

    if not configured:
        return []

    return resolve_column_list(
        df,
        configured,
        label="reporting value column",
    )


def resolve_french_labels(report_cfg: dict[str, Any]) -> dict[str, str]:
    for key in ("french_labels", "labels", "column_labels"):
        value = report_cfg.get(key)
        if isinstance(value, dict):
            return {str(k): str(v) for k, v in value.items()}
    return {}


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------


def main() -> None:
    print("=" * 80)
    print("LOCAL REMARKETING PIPELINE TEST")
    print("=" * 80)

    configs = load_configs()
    workbook_cfg = configs["workbook"]
    business_cfg = configs["business"]
    mapping_cfg = configs["columns"]
    report_cfg = configs["report"]

    input_file = resolve_input_file()
    output_file = OUTPUT_DIR / "remarketing_report_local_test.xlsx"

    print(f"\nProject root : {ROOT}")
    print(f"Input file   : {input_file}")
    print(f"Output file  : {output_file}")
    print(f"Input size   : {input_file.stat().st_size / 1024 / 1024:.2f} MB")

    # ------------------------------------------------------------------
    # 1/10 - workbook validation
    # ------------------------------------------------------------------

    print("\n[1/10] Validating workbook...")

    configured_required_sheets = workbook_cfg.get("required_sheets", []) or []
    if isinstance(configured_required_sheets, str):
        configured_required_sheets = [configured_required_sheets]

    required_sheets = [
        resolve_sheet_name(input_file, str(sheet))
        for sheet in configured_required_sheets
    ]

    max_bytes = int(workbook_cfg.get("max_bytes", 250_000_000))

    validation = validate_workbook(
        input_file,
        required_sheets=required_sheets,
        max_bytes=max_bytes,
    )

    print("Workbook validation passed.")
    print(f"Validation result: {validation}")

    # ------------------------------------------------------------------
    # 2/10 - read source worksheet
    # ------------------------------------------------------------------

    print("\n[2/10] Reading configured worksheets...")

    configured_source_sheet = workbook_cfg.get("source_sheet", "Remarketing")
    source_sheet = resolve_sheet_name(input_file, str(configured_source_sheet))

    data = read_configured(input_file, [source_sheet])
    working = data[source_sheet].copy()
    source = working.copy()

    print(
        f"  - {source_sheet}: "
        f"{working.shape[0]:,} rows x {working.shape[1]:,} columns"
    )
    print(f"\nPrimary processing sheet: {source_sheet}")

    # ------------------------------------------------------------------
    # 3/10 - source schema validation
    # ------------------------------------------------------------------

    print("\n[3/10] Validating required SOURCE columns...")

    validate_required_source_columns(
        working,
        workbook_cfg,
        business_cfg,
    )

    # Resolve config source names to the real dataframe names once.
    runtime_business_cfg = resolve_runtime_business_config(
        working,
        business_cfg,
    )

    # ------------------------------------------------------------------
    # 4/10 - standardize
    # ------------------------------------------------------------------

    print("\n[4/10] Standardizing data...")

    text_columns = resolve_column_list(
        working,
        workbook_cfg.get("text_columns", []) or [],
        label="text column",
    )

    working = standardize(
        working,
        text_columns=text_columns,
    )

    print(
        f"Standardization completed for {len(text_columns)} text column(s)."
    )

    # ------------------------------------------------------------------
    # 5/10 - lookups
    # ------------------------------------------------------------------

    print("\n[5/10] Applying configured reference lookups...")

    working = apply_reference_lookups(
        working,
        input_file,
        mapping_cfg,
    )

    print("Reference lookup stage completed.")

    # Re-resolve runtime source names after lookup/standardization in case
    # dataframe columns were preserved or augmented.
    runtime_business_cfg = resolve_runtime_business_config(
        working,
        business_cfg,
    )

    # ------------------------------------------------------------------
    # Resolve period
    # ------------------------------------------------------------------

    period = resolve_period(working, runtime_business_cfg)
    print(f"\nReporting period: {period}")

    # ------------------------------------------------------------------
    # 6/10 - business rules
    # ------------------------------------------------------------------

    print("\n[6/10] Applying business calculation rules...")

    working = apply_rules(
        working,
        runtime_business_cfg,
        period,
    )

    validate_derived_columns(
        working,
        runtime_business_cfg,
    )

    print("Business calculation rules completed.")

    # ------------------------------------------------------------------
    # 7/10 - monthly metrics
    # ------------------------------------------------------------------

    print("\n[7/10] Building monthly metrics...")

    columns_cfg = runtime_business_cfg.get("columns", {}) or {}
    event_date = columns_cfg.get("event_date")

    if not event_date:
        raise KeyError(
            "business_rules.yaml must configure columns.event_date."
        )

    event_date = resolve_actual_column(
        working,
        str(event_date),
        label="monthly event_date",
    )

    value_columns = resolve_value_columns(
        working,
        report_cfg,
    )

    monthly = monthly_metrics(
        working,
        date_col=event_date,
        value_cols=value_columns,
        period=period,
    )

    print(
        f"Monthly metrics created: "
        f"{monthly.shape[0]:,} rows x {monthly.shape[1]:,} columns"
    )

    # ------------------------------------------------------------------
    # 8/10 - YTD
    # ------------------------------------------------------------------

    print("\n[8/10] Building YTD dataset...")

    ytd = ytd_slice(
        working,
        date_col=event_date,
        period=period,
    )

    print(
        f"YTD dataset created: "
        f"{ytd.shape[0]:,} rows x {ytd.shape[1]:,} columns"
    )

    # ------------------------------------------------------------------
    # 9/10 - reconciliation
    # ------------------------------------------------------------------

    print("\n[9/10] Reconciling results...")

    excluded = source.iloc[0:0].copy()

    reconciliation = reconcile(
        source=source,
        valid=working,
        excluded=excluded,
        monthly=monthly,
        ytd=ytd,
    )

    print("Reconciliation result:")
    print(reconciliation)

    # ------------------------------------------------------------------
    # 10/10 - report
    # ------------------------------------------------------------------

    print("\n[10/10] Building local Excel report...")

    output_cfg = workbook_cfg.get("output", {}) or {}

    detail_sheet = output_cfg.get("detail_sheet", "Remarketing")
    remarketing_sheet = output_cfg.get(
        "remarketing_sheet",
        "Rapport Remarketing",
    )
    era_sheet = output_cfg.get("era_sheet", "Rapport ERA")

    report_sheets = {
        str(detail_sheet): working,
        str(remarketing_sheet): monthly,
        str(era_sheet): ytd,
    }

    french_labels = resolve_french_labels(report_cfg)

    build_report(
        output_file,
        sheets=report_sheets,
        french_labels=french_labels,
    )

    if not output_file.exists():
        raise RuntimeError(
            f"Report builder returned without creating {output_file}"
        )

    print(f"Report created:\n{output_file}")

    print("\n" + "=" * 80)
    print("LOCAL PIPELINE EXECUTION COMPLETED")
    print("=" * 80)
    print(
        "\nImportant: technical execution success does not prove the "
        "business calculations are correct. Reconcile the generated "
        "report against the analyst's manual Excel output before accepting "
        "the business logic."
    )


if __name__ == "__main__":
    main()
