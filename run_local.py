from __future__ import annotations

import os
import re
import sys
import time
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

# -----------------------------------------------------------------------------
# Make project imports work whether this file is run as:
#   python scripts/run_local.py
# or:
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
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data if isinstance(data, dict) else {}


def recursive_get(obj: Any, *keys: str) -> Any:
    """Find the first matching key anywhere inside a nested dict/list."""
    wanted = {k.lower() for k in keys}

    if isinstance(obj, dict):
        for key, value in obj.items():
            if str(key).lower() in wanted:
                return value
        for value in obj.values():
            found = recursive_get(value, *keys)
            if found is not None:
                return found

    if isinstance(obj, list):
        for value in obj:
            found = recursive_get(value, *keys)
            if found is not None:
                return found

    return None


def find_input_workbook() -> Path:
    """
    Select the local workbook copy.

    Optional override:
        $env:LOCAL_INPUT_FILE="C:\\path\\file.xlsx"
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
        p
        for p in INPUT_DIR.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".xlsm"}
        and not p.name.startswith("~$")
    )

    if not candidates:
        raise FileNotFoundError(
            f"No .xlsx/.xlsm workbook found in: {INPUT_DIR}"
        )

    if len(candidates) > 1:
        names = "\n".join(f"  - {p.name}" for p in candidates)
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
        return [str(x) for x in value]
    if isinstance(value, dict):
        return [str(k) for k in value.keys()]
    return []


def resolve_required_sheets(
    workbook_cfg: dict[str, Any], available: list[str]
) -> list[str]:
    configured = recursive_get(
        workbook_cfg,
        "required_sheets",
        "input_sheets",
        "source_sheets",
        "sheets",
    )
    requested = as_string_list(configured)

    # If the config does not explicitly define sheets, validate the workbook
    # itself and load all sheets. This is local testing only.
    if not requested:
        return available

    available_lookup = {s.casefold(): s for s in available}
    resolved: list[str] = []

    for sheet in requested:
        match = available_lookup.get(sheet.casefold())
        if match:
            resolved.append(match)

    return list(dict.fromkeys(resolved))


def resolve_primary_sheet(
    workbook_cfg: dict[str, Any],
    frames: dict[str, pd.DataFrame],
) -> str:
    configured = recursive_get(
        workbook_cfg,
        "primary_sheet",
        "main_sheet",
        "source_sheet",
        "data_sheet",
    )

    if isinstance(configured, str):
        for name in frames:
            if name.casefold() == configured.casefold():
                return name

    if not frames:
        raise RuntimeError("No worksheets were loaded.")

    # Safe fallback for local testing: use the loaded sheet with most rows.
    # We print the selection so it is never silent.
    return max(frames, key=lambda name: len(frames[name]))


def extract_column_mapping(
    cfg: dict[str, Any], sheet_name: str
) -> dict[str, str]:
    """Support flat or sheet-specific source->canonical mappings."""
    if not cfg:
        return {}

    candidate: Any = cfg

    for key in ("columns", "column_mapping", "mappings"):
        value = cfg.get(key)
        if isinstance(value, dict):
            candidate = value
            break

    # Sheet-specific mapping
    if isinstance(candidate, dict):
        for key, value in candidate.items():
            if str(key).casefold() == sheet_name.casefold() and isinstance(value, dict):
                candidate = value
                break

    if not isinstance(candidate, dict):
        return {}

    result: dict[str, str] = {}
    for source, target in candidate.items():
        if isinstance(source, str) and isinstance(target, str):
            result[source] = target

    return result


def rename_columns_if_configured(
    df: pd.DataFrame,
    mapping_cfg: dict[str, Any],
    sheet_name: str,
) -> pd.DataFrame:
    mapping = extract_column_mapping(mapping_cfg, sheet_name)
    if not mapping:
        return df

    existing = {str(c).casefold(): c for c in df.columns}
    rename_map: dict[Any, str] = {}

    for source, target in mapping.items():
        actual = existing.get(source.casefold())
        if actual is not None:
            rename_map[actual] = target

    return df.rename(columns=rename_map) if rename_map else df


def resolve_text_columns(
    workbook_cfg: dict[str, Any],
    df: pd.DataFrame,
) -> list[str]:
    configured = recursive_get(
        workbook_cfg,
        "text_columns",
        "string_columns",
    )
    columns = [c for c in as_string_list(configured) if c in df.columns]
    if columns:
        return columns

    # Local fallback: standardize object/string columns only.
    return [
        str(col)
        for col in df.columns
        if pd.api.types.is_object_dtype(df[col])
        or pd.api.types.is_string_dtype(df[col])
    ]


def apply_configured_lookups(
    df: pd.DataFrame,
    frames: dict[str, pd.DataFrame],
    workbook_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
) -> pd.DataFrame:
    lookups = recursive_get(business_cfg, "lookups", "reference_lookups")
    if lookups is None:
        lookups = recursive_get(workbook_cfg, "lookups", "reference_lookups")

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

        if not all([ref_sheet, left_key, right_key, value_columns]):
            print(f"  - Lookup {index}: skipped (incomplete configuration)")
            continue

        matching_sheet = next(
            (s for s in frames if s.casefold() == str(ref_sheet).casefold()),
            None,
        )
        if matching_sheet is None:
            if required:
                raise KeyError(f"Required lookup sheet not loaded: {ref_sheet}")
            print(f"  - Lookup {index}: reference sheet not loaded: {ref_sheet}")
            continue

        values = as_string_list(value_columns)
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


def resolve_column_name(
    configs: list[dict[str, Any]],
    df: pd.DataFrame,
    config_keys: tuple[str, ...],
    fallback_patterns: tuple[str, ...],
    label: str,
    required: bool = True,
) -> str | None:
    for cfg in configs:
        configured = recursive_get(cfg, *config_keys)
        if isinstance(configured, str):
            for column in df.columns:
                if str(column).casefold() == configured.casefold():
                    return str(column)

    # Conservative fallback: only accept a unique matching column.
    matches: list[str] = []
    for column in df.columns:
        normalized = re.sub(r"[^a-z0-9]+", " ", str(column).casefold()).strip()
        if any(pattern in normalized for pattern in fallback_patterns):
            matches.append(str(column))

    matches = list(dict.fromkeys(matches))

    if len(matches) == 1:
        print(f"  - Auto-detected {label}: {matches[0]}")
        return matches[0]

    if required:
        raise KeyError(
            f"Could not safely determine {label}. "
            f"Configure one of {config_keys}. Candidates: {matches or 'none'}"
        )

    return None


def resolve_period(
    configs: list[dict[str, Any]],
    df: pd.DataFrame,
    date_col: str,
) -> pd.Period:
    # Explicit environment override is easiest for local regression testing.
    override = os.getenv("TEST_PERIOD")
    if override:
        return pd.Period(override, freq="M")

    for cfg in configs:
        value = recursive_get(cfg, "period", "reporting_period", "test_period")
        if value is not None:
            try:
                return pd.Period(str(value), freq="M")
            except Exception:
                pass

    dates = pd.to_datetime(df[date_col], errors="coerce").dropna()
    if dates.empty:
        raise ValueError(
            f"Cannot infer reporting period because '{date_col}' contains no valid dates. "
            "Set TEST_PERIOD, e.g. $env:TEST_PERIOD='2026-07'."
        )

    inferred = dates.max().to_period("M")
    print(f"  - TEST_PERIOD not set; inferred reporting period: {inferred}")
    return inferred


def resolve_value_columns(
    report_cfg: dict[str, Any],
    business_cfg: dict[str, Any],
    df: pd.DataFrame,
) -> list[str]:
    for cfg in (report_cfg, business_cfg):
        value = recursive_get(
            cfg,
            "value_columns",
            "metric_columns",
            "monthly_value_columns",
        )
        columns = [c for c in as_string_list(value) if c in df.columns]
        if columns:
            return columns

    # Conservative local fallback: numeric columns only.
    return [
        str(c)
        for c in df.select_dtypes(include="number").columns
    ]


def resolve_french_labels(
    report_cfg: dict[str, Any],
    mapping_cfg: dict[str, Any],
    primary_sheet: str,
) -> dict[str, str]:
    labels = recursive_get(report_cfg, "french_labels", "output_labels", "labels")
    if isinstance(labels, dict):
        return {
            str(k): str(v)
            for k, v in labels.items()
            if isinstance(k, str) and isinstance(v, str)
        }

    # If column_mapping is source(French)->canonical, reverse it for output.
    source_to_canonical = extract_column_mapping(mapping_cfg, primary_sheet)
    return {canonical: source for source, canonical in source_to_canonical.items()}


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
    for idx, name in enumerate(available_sheets, start=1):
        print(f"  {idx:>2}. {name}")

    required_sheets = resolve_required_sheets(workbook_cfg, available_sheets)

    if not required_sheets:
        raise RuntimeError(
            "None of the configured sheets exist in the workbook. "
            "Check config/workbook_config.yaml."
        )

    print("\nSheets selected for local processing:")
    for name in required_sheets:
        print(f"  - {name}")

    # ------------------------------------------------------------------
    # 2. Workbook validation
    # ------------------------------------------------------------------
    print("\n[2/10] Validating workbook...")

    max_bytes = recursive_get(workbook_cfg, "max_bytes", "max_file_bytes")
    if not isinstance(max_bytes, int):
        # Local testing default: 1 GiB. This is only a validation ceiling.
        max_bytes = 1024 * 1024 * 1024

    validation = validate_workbook(
        input_file,
        required_sheets,
        max_bytes,
    )

    print("Workbook validation passed.")
    if validation:
        print(f"Validation result: {validation}")

    # ------------------------------------------------------------------
    # 3. Load configured sheets
    # ------------------------------------------------------------------
    print("\n[3/10] Reading configured worksheets...")
    frames = read_configured(input_file, required_sheets)

    for name, df in frames.items():
        print_frame_summary(name, df)

    primary_sheet = resolve_primary_sheet(workbook_cfg, frames)
    print(f"\nPrimary processing sheet: {primary_sheet}")

    # ------------------------------------------------------------------
    # 4. Rename configured source columns + standardize text values
    # ------------------------------------------------------------------
    print("\n[4/10] Standardizing data...")

    standardized_frames: dict[str, pd.DataFrame] = {}

    for name, source_df in frames.items():
        df = source_df.copy()
        df = rename_columns_if_configured(df, mapping_cfg, name)
        text_columns = resolve_text_columns(workbook_cfg, df)

        if text_columns:
            df = standardize(df, text_columns)

        standardized_frames[name] = df
        print(
            f"  - {name}: standardized "
            f"{len(text_columns)} text column(s)"
        )

    working = standardized_frames[primary_sheet].copy()

    # ------------------------------------------------------------------
    # 5. Reference lookups / mappings
    # ------------------------------------------------------------------
    print("\n[5/10] Applying configured reference lookups...")
    working = apply_configured_lookups(
        working,
        standardized_frames,
        workbook_cfg,
        business_cfg,
    )
    print("Reference lookup stage completed.")

    # ------------------------------------------------------------------
    # 6. Business calculations
    # ------------------------------------------------------------------
    print("\n[6/10] Applying business calculation rules...")

    configs = [business_cfg, workbook_cfg, report_cfg]

    date_col = resolve_column_name(
        configs,
        working,
        ("date_col", "date_column", "report_date_column", "transaction_date_column"),
        ("date",),
        "date column",
    )

    month_col = resolve_column_name(
        configs,
        working,
        ("month_col", "month_column"),
        ("month", "mois"),
        "month column",
        required=False,
    )

    year_col = resolve_column_name(
        configs,
        working,
        ("year_col", "year_column"),
        ("year", "annee", "année"),
        "year column",
        required=False,
    )

    # apply_rules requires names for month/year. If they do not already exist,
    # create local canonical columns from the configured date column.
    parsed_dates = pd.to_datetime(working[date_col], errors="coerce")

    if month_col is None:
        month_col = "__report_month"
        working[month_col] = parsed_dates.dt.month
        print(f"  - Created local month column: {month_col}")

    if year_col is None:
        year_col = "__report_year"
        working[year_col] = parsed_dates.dt.year
        print(f"  - Created local year column: {year_col}")

    period = resolve_period(configs, working, date_col)
    print(f"  - Processing period: {period}")

    calculated = apply_rules(
        working,
        business_cfg,
        date_col,
        month_col,
        year_col,
    )

    if calculated is None:
        raise RuntimeError("apply_rules() returned None; expected a DataFrame.")

    print_frame_summary("Calculated data", calculated)

    # ------------------------------------------------------------------
    # 7. Monthly reporting
    # ------------------------------------------------------------------
    print("\n[7/10] Calculating monthly metrics...")

    value_columns = resolve_value_columns(report_cfg, business_cfg, calculated)
    print(f"  - Metric/value columns: {value_columns}")

    monthly = monthly_metrics(
        calculated,
        date_col,
        value_columns,
        period,
    )

    print_frame_summary("Monthly report", monthly)

    # ------------------------------------------------------------------
    # 8. YTD reporting
    # ------------------------------------------------------------------
    print("\n[8/10] Calculating YTD dataset...")

    ytd = ytd_slice(
        calculated,
        date_col,
        period,
    )

    print_frame_summary("YTD dataset", ytd)

    # ------------------------------------------------------------------
    # 9. Reconciliation
    # ------------------------------------------------------------------
    print("\n[9/10] Running reconciliation...")

    source_df = frames[primary_sheet]
    source_count = len(source_df)
    valid_count = len(calculated)
    excluded_count = max(source_count - valid_count, 0)

    reconciliation = reconcile(
        source_count,
        valid_count,
        excluded_count,
        monthly=monthly,
        ytd=ytd,
    )

    print("Reconciliation result:")
    print(reconciliation)

    # ------------------------------------------------------------------
    # 10. Build local report workbook
    # ------------------------------------------------------------------
    print("\n[10/10] Building output workbook...")

    french_labels = resolve_french_labels(
        report_cfg,
        mapping_cfg,
        primary_sheet,
    )

    output_sheets: dict[str, pd.DataFrame] = {
        "Processed_Data": calculated,
        "Monthly_Report": monthly,
        "YTD_Data": ytd,
    }

    build_report(
        output_file,
        output_sheets,
        french_labels,
    )

    if not output_file.exists():
        raise RuntimeError(
            f"build_report() completed but output file was not created: {output_file}"
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
