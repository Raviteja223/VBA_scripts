from __future__ import annotations

import os
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
# COLUMN VALIDATION
# -----------------------------------------------------------------------------


def validate_required_source_columns(
    df: pd.DataFrame,
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> None:
    """
    Validate only columns that must already exist in the source workbook.

    Derived fields such as 'Statut calculé' and age buckets are intentionally
    excluded because calculations.apply_rules() is expected to create them.
    """

    required: set[str] = set()

    for key in (
        "required_columns",
        "business_keys",
        "text_columns",
    ):
        values = workbook_cfg.get(key, []) or []

        if isinstance(values, str):
            values = [values]

        for column in values:
            if column:
                required.add(str(column))

    columns_cfg = business_cfg.get("columns", {}) or {}

    for logical_key in BUSINESS_SOURCE_COLUMN_KEYS:
        column = columns_cfg.get(logical_key)
        if column:
            required.add(str(column))

    missing = sorted(
        column
        for column in required
        if column not in df.columns
    )

    if missing:
        available = "\n".join(
            f"  - {column!r}"
            for column in df.columns
        )

        raise KeyError(
            "Required SOURCE column(s) missing from the processing sheet: "
            + ", ".join(missing)
            + "\n\nAvailable source columns:\n"
            + available
        )

    print("Required source-column validation passed.")
    print("Validated source columns:")

    for column in sorted(required):
        print(f"  - {column}")



def validate_derived_columns(
    df: pd.DataFrame,
    business_cfg: dict[str, Any],
) -> None:
    """Validate fields that apply_rules() is expected to generate."""

    columns_cfg = business_cfg.get("columns", {}) or {}

    expected = [
        str(columns_cfg[key])
        for key in BUSINESS_DERIVED_COLUMN_KEYS
        if columns_cfg.get(key)
    ]

    missing = [
        column
        for column in expected
        if column not in df.columns
    ]

    if missing:
        raise KeyError(
            "Calculation step did not create expected derived column(s): "
            + ", ".join(missing)
        )

    print("Derived-column validation passed.")
    print("Generated columns:")

    for column in expected:
        print(f"  - {column}")


# -----------------------------------------------------------------------------
# PERIOD
# -----------------------------------------------------------------------------


def resolve_period(
    df: pd.DataFrame,
    business_cfg: dict[str, Any],
) -> str:
    """
    Resolve YYYY-MM for local execution.

    TEST_PERIOD is preferred. If absent, infer from the configured event_date.
    """

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

    if event_date not in df.columns:
        raise KeyError(
            f"Configured event_date column {event_date!r} was not found."
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
    """
    Apply explicitly configured workbook lookups only.

    If column_mapping.yaml contains no supported lookup list, this is a no-op.
    """

    result = df
    definitions = _lookup_definitions(mapping_cfg)

    if not definitions:
        print("No configured reference lookups found; skipping lookup stage.")
        return result

    available_sheets = set(sheet_names(input_file))

    for index, definition in enumerate(definitions, start=1):
        sheet = definition.get("sheet") or definition.get("reference_sheet")
        left_key = definition.get("left_key")
        right_key = definition.get("right_key")
        value_columns = (
            definition.get("value_columns")
            or definition.get("columns")
            or []
        )
        required = bool(definition.get("required", False))

        if isinstance(value_columns, str):
            value_columns = [value_columns]

        missing_config = [
            name
            for name, value in {
                "sheet": sheet,
                "left_key": left_key,
                "right_key": right_key,
            }.items()
            if not value
        ]

        if missing_config:
            raise ValueError(
                f"Lookup #{index} is missing config field(s): "
                + ", ".join(missing_config)
            )

        if sheet not in available_sheets:
            raise KeyError(
                f"Lookup #{index} reference sheet {sheet!r} does not exist."
            )

        if left_key not in result.columns:
            raise KeyError(
                f"Lookup #{index} source key {left_key!r} does not exist "
                "in the processing dataframe."
            )

        reference = read_sheet(input_file, sheet)

        required_reference_columns = [right_key, *value_columns]
        missing_reference_columns = [
            column
            for column in required_reference_columns
            if column not in reference.columns
        ]

        if missing_reference_columns:
            raise KeyError(
                f"Lookup #{index} reference sheet {sheet!r} is missing: "
                + ", ".join(missing_reference_columns)
            )

        result = left_lookup(
            result,
            reference,
            left_key=left_key,
            right_key=right_key,
            value_columns=value_columns,
            required=required,
        )

        print(
            f"Applied lookup #{index}: {sheet!r} "
            f"({left_key!r} -> {right_key!r})"
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

    if configured:
        missing = [c for c in configured if c not in df.columns]

        if missing:
            raise KeyError(
                "Configured reporting value column(s) missing: "
                + ", ".join(missing)
            )

        return configured

    # Do not invent business metrics. An empty list lets the project's
    # monthly_reporting implementation decide whether count-only reporting
    # is supported. If it is not, it will fail clearly at the reporting stage.
    return []


def resolve_french_labels(report_cfg: dict[str, Any]) -> dict[str, str]:
    for key in (
        "french_labels",
        "labels",
        "column_labels",
    ):
        value = report_cfg.get(key)
        if isinstance(value, dict):
            return {
                str(k): str(v)
                for k, v in value.items()
            }
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

    required_sheets = workbook_cfg.get("required_sheets", []) or []
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

    source_sheet = workbook_cfg.get("source_sheet", "Remarketing")
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

    # ------------------------------------------------------------------
    # 4/10 - standardize
    # ------------------------------------------------------------------

    print("\n[4/10] Standardizing data...")

    text_columns = workbook_cfg.get("text_columns", []) or []
    if isinstance(text_columns, str):
        text_columns = [text_columns]

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

    # ------------------------------------------------------------------
    # Resolve period once source data is available
    # ------------------------------------------------------------------

    period = resolve_period(working, business_cfg)
    print(f"\nReporting period: {period}")

    # ------------------------------------------------------------------
    # 6/10 - business rules
    # ------------------------------------------------------------------

    print("\n[6/10] Applying business calculation rules...")

    working = apply_rules(
        working,
        business_cfg,
        period,
    )

    validate_derived_columns(
        working,
        business_cfg,
    )

    print("Business calculation rules completed.")

    # ------------------------------------------------------------------
    # 7/10 - monthly metrics
    # ------------------------------------------------------------------

    print("\n[7/10] Building monthly metrics...")

    columns_cfg = business_cfg.get("columns", {}) or {}
    event_date = columns_cfg.get("event_date")

    if not event_date:
        raise KeyError(
            "business_rules.yaml must configure columns.event_date."
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

    detail_sheet = output_cfg.get(
        "detail_sheet",
        "Remarketing",
    )

    remarketing_sheet = output_cfg.get(
        "remarketing_sheet",
        "Rapport Remarketing",
    )

    era_sheet = output_cfg.get(
        "era_sheet",
        "Rapport ERA",
    )

    report_sheets = {
        detail_sheet: working,
        remarketing_sheet: monthly,
        era_sheet: ytd,
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
        "\nImportant: execution success only proves the technical flow. "
        "The generated numbers still need reconciliation against the "
        "analyst's manual Excel report before the business logic is accepted."
    )


if __name__ == "__main__":
    main()
