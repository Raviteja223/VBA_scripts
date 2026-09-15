from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

# -----------------------------------------------------------------------------
# Make project imports work for both:
#   python scripts/run_local.py
#   python -m scripts.run_local
# -----------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from include.remarketing.excel_reader import read_configured, sheet_names
from include.remarketing.workbook_validation import validate_workbook
from include.remarketing.standardization import standardize
from include.remarketing.mappings import left_lookup
from include.remarketing.calculations import apply_rules
from include.remarketing.monthly_reporting import monthly_metrics
from include.remarketing.ytd_reporting import ytd_slice
from include.remarketing.report_builder import build_report
from include.remarketing.reconciliation import reconcile


# -----------------------------------------------------------------------------
# Local paths
# -----------------------------------------------------------------------------
INPUT_DIR = ROOT / "local_test" / "input"
OUTPUT_DIR = ROOT / "local_test" / "output"

WORKBOOK_CONFIG_PATH = ROOT / "config" / "workbook_config.yaml"
BUSINESS_RULES_PATH = ROOT / "config" / "business_rules.yaml"
COLUMN_MAPPING_PATH = ROOT / "config" / "column_mapping.yaml"
REPORT_CONFIG_PATH = ROOT / "config" / "report_config.yaml"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle)

    return data if isinstance(data, dict) else {}


def find_input_workbook() -> Path:
    """Find the single workbook under local_test/input.

    Optional override in PowerShell:
        $env:LOCAL_INPUT_FILE="C:\\path\\to\\file.xlsx"
    """
    override = os.getenv("LOCAL_INPUT_FILE")
    if override:
        path = Path(override).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"LOCAL_INPUT_FILE does not exist: {path}")
        return path

    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Input directory does not exist: {INPUT_DIR}")

    candidates = sorted(
        path
        for path in INPUT_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
    )

    if not candidates:
        raise FileNotFoundError(
            f"No .xlsx/.xlsm workbook found in: {INPUT_DIR}"
        )

    if len(candidates) > 1:
        names = "\n".join(f"  - {path.name}" for path in candidates)
        raise RuntimeError(
            "More than one workbook exists in local_test/input. "
            "Keep only the workbook you want to test, or set LOCAL_INPUT_FILE.\n"
            f"Found:\n{names}"
        )

    return candidates[0]


def as_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, (list, tuple, set)):
        return [str(item) for item in value]
    return []


def get_source_sheet(workbook_cfg: dict[str, Any], available: list[str]) -> str:
    configured = workbook_cfg.get("source_sheet") or workbook_cfg.get("primary_sheet")

    if isinstance(configured, str):
        for sheet in available:
            if sheet.casefold() == configured.casefold():
                return sheet
        raise KeyError(
            f"Configured source_sheet '{configured}' was not found in the workbook."
        )

    if not available:
        raise RuntimeError("Workbook has no worksheets.")

    return available[0]


def get_required_sheets(
    workbook_cfg: dict[str, Any],
    available: list[str],
    source_sheet: str,
) -> list[str]:
    requested = as_string_list(workbook_cfg.get("required_sheets"))

    if not requested:
        requested = [source_sheet]

    available_lookup = {sheet.casefold(): sheet for sheet in available}
    resolved: list[str] = []
    missing: list[str] = []

    for sheet in requested:
        actual = available_lookup.get(sheet.casefold())
        if actual is None:
            missing.append(sheet)
        else:
            resolved.append(actual)

    if missing:
        raise KeyError(
            "Required worksheet(s) missing: " + ", ".join(missing)
        )

    if source_sheet not in resolved:
        resolved.insert(0, source_sheet)

    return list(dict.fromkeys(resolved))


def validate_required_columns(
    df: pd.DataFrame,
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> None:
    required = as_string_list(workbook_cfg.get("required_columns"))

    columns_cfg = business_cfg.get("columns")
    if isinstance(columns_cfg, dict):
        # These are source columns that the calculation code actually reads.
        for key in (
            "status_raw",
            "event_date",
            "settlement_date",
            "return_date",
        ):
            value = columns_cfg.get(key)
            if isinstance(value, str):
                required.append(value)

    required = list(dict.fromkeys(required))
    missing = [column for column in required if column not in df.columns]

    if missing:
        raise KeyError(
            "Required source column(s) missing from the processing sheet: "
            + ", ".join(missing)
        )


def get_event_date_column(
    business_cfg: dict[str, Any],
    df: pd.DataFrame,
) -> str:
    """Use the explicit event date configured in business_rules.yaml.

    The workbook contains several valid date columns, so this runner deliberately
    does not guess which one should drive monthly/YTD reporting.
    """
    columns_cfg = business_cfg.get("columns")
    if not isinstance(columns_cfg, dict):
        raise KeyError("business_rules.yaml is missing the 'columns' section.")

    event_date = columns_cfg.get("event_date")
    if not isinstance(event_date, str) or not event_date.strip():
        raise KeyError(
            "business_rules.yaml is missing columns.event_date."
        )

    if event_date not in df.columns:
        raise KeyError(
            f"Configured event date column '{event_date}' was not found "
            "in the processing sheet."
        )

    return event_date


def resolve_period(df: pd.DataFrame, event_date_col: str) -> str:
    """Return reporting period as YYYY-MM.

    Preferred local override:
        $env:TEST_PERIOD="2026-07"
    """
    override = os.getenv("TEST_PERIOD")
    if override:
        try:
            return str(pd.Period(override, freq="M"))
        except Exception as exc:
            raise ValueError(
                f"Invalid TEST_PERIOD '{override}'. Expected format like 2026-07."
            ) from exc

    dates = pd.to_datetime(df[event_date_col], errors="coerce").dropna()
    if dates.empty:
        raise ValueError(
            f"Cannot infer reporting period because '{event_date_col}' has no "
            "valid dates. Set TEST_PERIOD explicitly, for example: "
            "$env:TEST_PERIOD='2026-07'"
        )

    inferred = str(dates.max().to_period("M"))
    print(f"  - TEST_PERIOD not set; inferred reporting period: {inferred}")
    return inferred


def resolve_text_columns(
    workbook_cfg: dict[str, Any],
    df: pd.DataFrame,
) -> list[str]:
    configured = as_string_list(workbook_cfg.get("text_columns"))
    if configured:
        missing = [column for column in configured if column not in df.columns]
        if missing:
            print(
                "  - Warning: configured text columns not found and skipped: "
                + ", ".join(missing)
            )
        return [column for column in configured if column in df.columns]

    return [
        str(column)
        for column in df.columns
        if pd.api.types.is_object_dtype(df[column])
        or pd.api.types.is_string_dtype(df[column])
    ]


def apply_configured_lookups(
    df: pd.DataFrame,
    frames: dict[str, pd.DataFrame],
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> pd.DataFrame:
    lookups = business_cfg.get("lookups") or workbook_cfg.get("lookups")
    if not isinstance(lookups, list):
        return df

    result = df

    for index, spec in enumerate(lookups, start=1):
        if not isinstance(spec, dict):
            continue

        ref_sheet = (
            spec.get("ref_sheet")
            or spec.get("reference_sheet")
            or spec.get("sheet")
        )
        left_key = spec.get("left_key")
        right_key = spec.get("right_key")
        value_columns = spec.get("value_columns") or spec.get("columns")
        required = bool(spec.get("required", False))

        values = as_string_list(value_columns)

        if not ref_sheet or not left_key or not right_key or not values:
            print(f"  - Lookup {index}: skipped (incomplete configuration)")
            continue

        matching_sheet = next(
            (
                name
                for name in frames
                if name.casefold() == str(ref_sheet).casefold()
            ),
            None,
        )

        if matching_sheet is None:
            if required:
                raise KeyError(f"Required lookup sheet not loaded: {ref_sheet}")
            print(f"  - Lookup {index}: reference sheet not loaded: {ref_sheet}")
            continue

        print(
            f"  - Lookup {index}: {left_key} -> "
            f"{matching_sheet}.{right_key} ({', '.join(values)})"
        )

        result = left_lookup(
            result,
            frames[matching_sheet],
            str(left_key),
            str(right_key),
            values,
            required=required,
        )

    return result


def get_value_columns(
    report_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
    df: pd.DataFrame,
) -> list[str]:
    for cfg in (report_cfg, business_cfg):
        for key in ("value_columns", "metric_columns", "monthly_value_columns"):
            configured = as_string_list(cfg.get(key))
            found = [column for column in configured if column in df.columns]
            if found:
                return found

    # Local fallback only: aggregate numeric columns if no explicit report metric
    # list exists yet.
    return [str(column) for column in df.select_dtypes(include="number").columns]


def get_french_labels(
    report_cfg: dict[str, Any],
    mapping_cfg: dict[str, Any],
) -> dict[str, str]:
    for key in ("french_labels", "output_labels", "labels"):
        labels = report_cfg.get(key)
        if isinstance(labels, dict):
            return {
                str(source): str(target)
                for source, target in labels.items()
                if isinstance(source, str) and isinstance(target, str)
            }

    # If the mapping file contains a canonical->French output mapping, use it.
    for key in ("french_labels", "output_labels"):
        labels = mapping_cfg.get(key)
        if isinstance(labels, dict):
            return {
                str(source): str(target)
                for source, target in labels.items()
                if isinstance(source, str) and isinstance(target, str)
            }

    return {}


def get_output_sheet_names(workbook_cfg: dict[str, Any]) -> tuple[str, str, str]:
    output_cfg = workbook_cfg.get("output")
    if not isinstance(output_cfg, dict):
        output_cfg = {}

    detail = str(output_cfg.get("detail_sheet") or "Processed_Data")
    monthly = str(output_cfg.get("remarketing_sheet") or "Monthly_Report")
    ytd = str(output_cfg.get("era_sheet") or "YTD_Data")

    return detail[:31], monthly[:31], ytd[:31]


def print_frame_summary(name: str, df: pd.DataFrame) -> None:
    print(f"  - {name}: {len(df):,} rows x {len(df.columns):,} columns")


# -----------------------------------------------------------------------------
# Main local pipeline
# -----------------------------------------------------------------------------
def main() -> None:
    started = time.perf_counter()

    print("=" * 88)
    print("LOCAL REMARKETING PIPELINE TEST")
    print("=" * 88)

    input_file = find_input_workbook()
    output_file = OUTPUT_DIR / "remarketing_report_local_test.xlsx"

    print(f"\nProject root : {ROOT}")
    print(f"Input file   : {input_file}")
    print(f"Output file  : {output_file}")
    print(f"Input size   : {input_file.stat().st_size / (1024 * 1024):.2f} MB")

    workbook_cfg = load_yaml(WORKBOOK_CONFIG_PATH)
    business_cfg = load_yaml(BUSINESS_RULES_PATH)
    mapping_cfg = load_yaml(COLUMN_MAPPING_PATH)
    report_cfg = load_yaml(REPORT_CONFIG_PATH)

    # ------------------------------------------------------------------
    # 1. Inspect workbook
    # ------------------------------------------------------------------
    print("\n[1/10] Inspecting workbook...")
    available_sheets = sheet_names(input_file)

    print(f"Found {len(available_sheets)} worksheet(s):")
    for index, name in enumerate(available_sheets, start=1):
        print(f"  {index:>2}. {name}")

    source_sheet = get_source_sheet(workbook_cfg, available_sheets)
    required_sheets = get_required_sheets(
        workbook_cfg,
        available_sheets,
        source_sheet,
    )

    print("\nSheets selected for local processing:")
    for name in required_sheets:
        print(f"  - {name}")

    # ------------------------------------------------------------------
    # 2. Workbook validation
    # ------------------------------------------------------------------
    print("\n[2/10] Validating workbook...")

    max_bytes = workbook_cfg.get("max_bytes", 1024 * 1024 * 1024)
    if not isinstance(max_bytes, int):
        max_bytes = int(max_bytes)

    validation = validate_workbook(
        input_file,
        required_sheets,
        max_bytes,
    )

    print("Workbook validation passed.")
    if validation:
        print(f"Validation result: {validation}")

    # ------------------------------------------------------------------
    # 3. Read configured worksheets
    # ------------------------------------------------------------------
    print("\n[3/10] Reading configured worksheets...")
    frames = read_configured(input_file, required_sheets)

    for name, frame in frames.items():
        print_frame_summary(name, frame)

    if source_sheet not in frames:
        raise KeyError(f"Source worksheet was not loaded: {source_sheet}")

    working = frames[source_sheet].copy()
    print(f"\nPrimary processing sheet: {source_sheet}")

    validate_required_columns(working, workbook_cfg, business_cfg)

    # ------------------------------------------------------------------
    # 4. Standardize configured text columns
    # ------------------------------------------------------------------
    print("\n[4/10] Standardizing data...")

    text_columns = resolve_text_columns(workbook_cfg, working)
    if text_columns:
        working = standardize(working, text_columns)

    print(
        f"  - {source_sheet}: standardized "
        f"{len(text_columns)} text column(s)"
    )

    # ------------------------------------------------------------------
    # 5. Reference lookups / mappings
    # ------------------------------------------------------------------
    print("\n[5/10] Applying configured reference lookups...")
    working = apply_configured_lookups(
        working,
        frames,
        workbook_cfg,
        business_cfg,
    )
    print("Reference lookup stage completed.")

    # ------------------------------------------------------------------
    # 6. Business calculations
    # ------------------------------------------------------------------
    print("\n[6/10] Applying business calculation rules...")

    event_date_col = get_event_date_column(business_cfg, working)
    period = resolve_period(working, event_date_col)

    print(f"  - Reporting/event date column: {event_date_col}")
    print(f"  - Processing period          : {period}")

    # Actual generated project API:
    #     apply_rules(df, cfg, period)
    calculated = apply_rules(
        working,
        business_cfg,
        period,
    )

    if calculated is None:
        raise RuntimeError("apply_rules() returned None; expected a DataFrame.")

    print_frame_summary("Calculated data", calculated)

    # ------------------------------------------------------------------
    # 7. Monthly reporting
    # ------------------------------------------------------------------
    print("\n[7/10] Calculating monthly metrics...")

    value_columns = get_value_columns(report_cfg, business_cfg, calculated)
    print(f"  - Metric/value columns: {value_columns}")

    monthly = monthly_metrics(
        calculated,
        event_date_col,
        value_columns,
        period,
    )

    if monthly is None:
        raise RuntimeError("monthly_metrics() returned None; expected a DataFrame.")

    print_frame_summary("Monthly report", monthly)

    # ------------------------------------------------------------------
    # 8. YTD reporting
    # ------------------------------------------------------------------
    print("\n[8/10] Calculating YTD dataset...")

    ytd = ytd_slice(
        calculated,
        event_date_col,
        period,
    )

    if ytd is None:
        raise RuntimeError("ytd_slice() returned None; expected a DataFrame.")

    print_frame_summary("YTD dataset", ytd)

    # ------------------------------------------------------------------
    # 9. Reconciliation
    # ------------------------------------------------------------------
    print("\n[9/10] Running reconciliation...")

    source_count = len(frames[source_sheet])
    valid_count = len(calculated)
    excluded_count = max(source_count - valid_count, 0)

    reconciliation_result = reconcile(
        source_count,
        valid_count,
        excluded_count,
        monthly=monthly,
        ytd=ytd,
    )

    print("Reconciliation result:")
    print(reconciliation_result)

    # ------------------------------------------------------------------
    # 10. Build output workbook
    # ------------------------------------------------------------------
    print("\n[10/10] Building output workbook...")

    french_labels = get_french_labels(report_cfg, mapping_cfg)
    detail_sheet, monthly_sheet, ytd_sheet = get_output_sheet_names(workbook_cfg)

    output_sheets: dict[str, pd.DataFrame] = {
        detail_sheet: calculated,
        monthly_sheet: monthly,
        ytd_sheet: ytd,
    }

    build_report(
        output_file,
        output_sheets,
        french_labels,
    )

    if not output_file.exists():
        raise RuntimeError(
            "build_report() completed but the expected output file was not created: "
            f"{output_file}"
        )

    elapsed = time.perf_counter() - started

    print("\n" + "=" * 88)
    print("LOCAL TEST COMPLETED SUCCESSFULLY")
    print("=" * 88)
    print(f"Output  : {output_file}")
    print(f"Duration: {elapsed:.2f} seconds")
    print(f"Period  : {period}")
    print("=" * 88)


if __name__ == "__main__":
    main()
