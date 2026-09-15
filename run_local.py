from pathlib import Path

from include.remarketing.excel_reader import read_workbook
from include.remarketing.workbook_validation import validate_workbook
from include.remarketing.standardization import standardize_data
from include.remarketing.mappings import apply_mappings
from include.remarketing.calculations import calculate_business_metrics
from include.remarketing.monthly_reporting import calculate_monthly_report
from include.remarketing.ytd_reporting import calculate_ytd_report
from include.remarketing.report_builder import build_reports
from include.remarketing.reconciliation import reconcile_reports


ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    ROOT
    / "local_test"
    / "input"
    / "REMARKETING_ET_REPOSSESSION_copy24072026.xlsx"
)

OUTPUT_DIR = ROOT / "local_test" / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def main():

    print("=" * 70)
    print("LOCAL REMARKETING TEST")
    print("=" * 70)

    print(f"\nInput file:")
    print(INPUT_FILE)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input workbook not found: {INPUT_FILE}"
        )

    # ---------------------------------------------------
    # STEP 1 - Validate workbook
    # ---------------------------------------------------

    print("\n[1/8] Validating workbook...")

    validate_workbook(INPUT_FILE)

    print("Workbook validation passed.")

    # ---------------------------------------------------
    # STEP 2 - Read workbook
    # ---------------------------------------------------

    print("\n[2/8] Reading workbook...")

    workbook_data = read_workbook(INPUT_FILE)

    print("Workbook loaded.")

    # ---------------------------------------------------
    # STEP 3 - Standardize
    # ---------------------------------------------------

    print("\n[3/8] Standardizing data...")

    standardized_data = standardize_data(
        workbook_data
    )

    print("Standardization completed.")

    # ---------------------------------------------------
    # STEP 4 - Mappings
    # ---------------------------------------------------

    print("\n[4/8] Applying mappings...")

    mapped_data = apply_mappings(
        standardized_data
    )

    print("Mappings completed.")

    # ---------------------------------------------------
    # STEP 5 - Business calculations
    # ---------------------------------------------------

    print("\n[5/8] Calculating business metrics...")

    calculated_data = calculate_business_metrics(
        mapped_data
    )

    print("Business calculations completed.")

    # ---------------------------------------------------
    # STEP 6 - Monthly + YTD
    # ---------------------------------------------------

    print("\n[6/8] Creating monthly/YTD reports...")

    monthly_report = calculate_monthly_report(
        calculated_data
    )

    ytd_report = calculate_ytd_report(
        calculated_data
    )

    print("Reporting calculations completed.")

    # ---------------------------------------------------
    # STEP 7 - Build Excel
    # ---------------------------------------------------

    print("\n[7/8] Building report workbook...")

    output_file = (
        OUTPUT_DIR
        / "remarketing_report_local_test.xlsx"
    )

    build_reports(
        calculated_data=calculated_data,
        monthly_report=monthly_report,
        ytd_report=ytd_report,
        output_file=output_file,
    )

    print(f"Generated:")
    print(output_file)

    # ---------------------------------------------------
    # STEP 8 - Reconciliation
    # ---------------------------------------------------

    print("\n[8/8] Running reconciliation...")

    reconciliation = reconcile_reports(
        calculated_data=calculated_data,
        monthly_report=monthly_report,
        ytd_report=ytd_report,
    )

    print("\nReconciliation results:")
    print(reconciliation)

    print("\n" + "=" * 70)
    print("LOCAL TEST COMPLETED")
    print("=" * 70)


if __name__ == "__main__":
    main()
