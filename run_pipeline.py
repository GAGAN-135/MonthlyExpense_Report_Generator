"""
One-click monthly expense pipeline.

Give it three inputs (expense sheet, attachments file, employee master) and it:
  1. Enriches the expense sheet with BSV + Region looked up from the employee
     master (a OneDrive-synced local copy of the online SharePoint sheet).
  2. Splits the enriched data by vertical -> region -> employee (Split Expenses.py).
  3. Downloads every receipt image into its employee's folder (Expense images download.py).

Only the CONFIG block below needs editing month to month.
"""

import importlib.util
import os
import re
from datetime import datetime
import sys

import pandas as pd

# ============================== CONFIG (edit me) ==============================
# 1) The monthly expense sheet (the one that is MISSING BSV / Region).
EXPENSE_FILE = r"C:\Users\Gagan.singh\Downloads\Expenses-Sep-26\Expense Report_0_2026-09-01 to 2026-09-30_55909.csv"

# 2) The attachments export (image CDN links).
ATTACHMENTS_FILE = r"C:\Users\Gagan.singh\Downloads\Expenses-Sep-26\Expense Attachment_2026-09-01_2026-09-30 (2).csv"

# 3) The employee master ("Channelier Employees.xlsx"), synced locally via OneDrive.
#    It lives on the analytics account's drive, so first get it SHARED with you,
#    then in OneDrive (browser) open "Shared" -> the file's folder -> "Add shortcut
#    to My files". It then syncs under your OneDrive; point this at that local copy.
EMPLOYEE_MASTER_FILE = r"C:\Users\Gagan.singh\Downloads\Expenses-Sep-26\Employees (3).xlsx"

# 4) Where the per-employee folders + images are written. A date-time stamp is
#    appended per run (e.g. "..._output_2026-07-11_1530"), so each run gets its
#    own clean folder and images never accumulate across runs.
OUTPUT_DIR = r"C:\Users\Gagan.singh\Downloads\Expenses-Sep-26"
# --- Column mapping (matched to "Channelier Employees.xlsx") -----------------
# Join key: the column that exists in BOTH the expense sheet and the master.
EXPENSE_KEY_COL = "employee code"      # expense sheet: EM366
MASTER_KEY_COL = "Employee ID"         # master: EM366

# The master columns that hold the vertical and the region.
MASTER_BSV_COL = "Business Sub Vertical"   # FS-CH / FS-KA / NUT / DIR (NOT "Business Vertical")
MASTER_REGION_COL = "Region"               # North-1, West, East, ...

# Employees with a valid BSV but a blank Region are placed here (matches the
# existing "0" folder behaviour). Set to None to instead skip them.
DEFAULT_REGION = "0"
VERTICALS = ["FS-CH", "NUT", "FS-KA", "DIR"]

# --- Combined overview sheet -------------------------------------------------
# A one-row-per-employee summary (name, id, BSV, region, total) written to its
# own folder inside the run output.
EMP_NAME_COL = "employee name"     # employee-name column in the expense sheet
TOTAL_AMOUNT_COL = "total amount"  # amount column to sum per employee
COMBINED_FOLDER_NAME = "Combined Expense"
# =============================================================================


def _load_sibling(module_name, filename):
    """Import a module that lives next to this file even if its name has spaces."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _normalize_key(value):
    """Normalize a join key (code or name) for reliable matching."""
    return re.sub(r"\s+", " ", str(value).replace("_", " ")).strip().lower()


def _read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path)
    raise ValueError(f"Unsupported file format: {ext}")


def enrich_with_bsv_region(expense_df, master_df):
    """Add 'BSV' and 'Region' columns to the expense data by looking them up in
    the employee master. Returns (enriched_df, unmatched_employee_names)."""
    for col, where in [(EXPENSE_KEY_COL, "expense sheet"),
                       (MASTER_KEY_COL, "employee master"),
                       (MASTER_BSV_COL, "employee master"),
                       (MASTER_REGION_COL, "employee master")]:
        src = expense_df if where == "expense sheet" else master_df
        if col not in src.columns:
            raise KeyError(
                f"Column '{col}' not found in the {where}. "
                f"Available columns: {list(src.columns)}"
            )

    # Build lookup: normalized key -> (BSV, Region). Last row wins on duplicates.
    lookup = {}
    for _, r in master_df.iterrows():
        lookup[_normalize_key(r[MASTER_KEY_COL])] = (r[MASTER_BSV_COL], r[MASTER_REGION_COL])

    keys = expense_df[EXPENSE_KEY_COL].map(_normalize_key)
    expense_df = expense_df.copy()
    expense_df["BSV"] = keys.map(lambda k: lookup.get(k, (None, None))[0])
    expense_df["Region"] = keys.map(lambda k: lookup.get(k, (None, None))[1])

    unmatched_mask = expense_df["BSV"].isna()
    name_col = "employee name" if "employee name" in expense_df.columns else EXPENSE_KEY_COL
    unmatched = sorted(expense_df.loc[unmatched_mask, name_col].astype(str).unique())

    # Keep employees whose vertical is known but region is blank (the "0" bucket).
    if DEFAULT_REGION is not None:
        expense_df["Region"] = expense_df["Region"].fillna(DEFAULT_REGION)

    return expense_df, unmatched


def build_combined_summary(enriched_df, out_dir):
    """Write a one-row-per-employee overview (name, id, BSV, region, total) to
    <out_dir>/<COMBINED_FOLDER_NAME>/Combined Expense Summary.xlsx. Returns its path."""
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side, Font
    from openpyxl.utils import get_column_letter

    summary = (
        enriched_df
        .groupby([EMP_NAME_COL, EXPENSE_KEY_COL, "BSV", "Region"], dropna=False, as_index=False)[TOTAL_AMOUNT_COL]
        .sum()
        .rename(columns={
            EMP_NAME_COL: "Employee Name",
            EXPENSE_KEY_COL: "Employee ID",
            "Region": "Sales Region",
            TOTAL_AMOUNT_COL: "Total Amount",
        })
    )
    summary = summary[["Employee Name", "Employee ID", "BSV", "Sales Region", "Total Amount"]]
    summary = summary.sort_values(["BSV", "Sales Region", "Total Amount"],
                                  ascending=[True, True, False]).reset_index(drop=True)
    summary["Total Amount"] = summary["Total Amount"].round(2)

    # Grand-total row at the bottom.
    grand_total = round(summary["Total Amount"].sum(), 2)
    summary.loc[len(summary)] = ["TOTAL", "", "", "", grand_total]

    combined_folder = os.path.join(out_dir, COMBINED_FOLDER_NAME)
    os.makedirs(combined_folder, exist_ok=True)
    out_path = os.path.join(combined_folder, "Combined Expense Summary.xlsx")
    summary.to_excel(out_path, index=False)

    # Match the per-employee files: thin borders, bold header + total, auto width.
    wb = load_workbook(out_path)
    ws = wb.active
    thin = Border(*(Side(style="thin"),) * 4)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = width + 2
    wb.save(out_path)
    return out_path, len(summary) - 1


def main():
    for label, path in [("Expense sheet", EXPENSE_FILE),
                        ("Attachments file", ATTACHMENTS_FILE),
                        ("Employee master", EMPLOYEE_MASTER_FILE)]:
        if not os.path.isfile(path):
            print(f"Error: {label} not found: '{path}'")
            sys.exit(1)

    if EMPLOYEE_MASTER_FILE.lower().endswith(".url"):
        print("Error: EMPLOYEE_MASTER_FILE points to a '.url' web shortcut, not the "
              "actual spreadsheet. 'Add shortcut to My files' on a single file only "
              "creates a link. Sync the containing FOLDER instead so the real "
              "'Channelier Employees.xlsx' lands on disk, then point this at that copy.")
        sys.exit(1)

    split_mod = _load_sibling("split_expenses", "Split Expenses.py")
    image_mod = _load_sibling("image_download", "Expense images download.py")

    # Each run writes to its own date-time-stamped folder so nothing accumulates.
    run_output_dir = f"{OUTPUT_DIR}_{datetime.now().strftime('%Y-%m-%d_%H%M')}"
    print(f"Output folder for this run: {run_output_dir}\n")

    # -- Step 1: enrich -------------------------------------------------------
    print("Step 1/4: Looking up BSV + Region from the employee master ...")
    expense_df = _read_table(EXPENSE_FILE)
    master_df = _read_table(EMPLOYEE_MASTER_FILE)
    enriched_df, unmatched = enrich_with_bsv_region(expense_df, master_df)
    print(f"  Enriched {len(enriched_df)} rows.")
    if unmatched:
        print(f"  WARNING: {len(unmatched)} employee(s) had no BSV/Region in the "
              f"master and will be SKIPPED by the split:")
        for name in unmatched:
            print(f"    - {name}")

    # -- Step 2: split --------------------------------------------------------
    print("\nStep 2/4: Splitting into per-employee files ...")
    split_mod.process_expenses(
        file_path=None,
        output_dir=run_output_dir,
        verticals_to_process=VERTICALS,
        second_filter="Region",
        df=enriched_df,
    )
    print("  Split complete.")

    # -- Step 3: combined overview -------------------------------------------
    print("\nStep 3/4: Building the combined overview sheet ...")
    combined_path, emp_count = build_combined_summary(enriched_df, run_output_dir)
    print(f"  {emp_count} employees summarised -> {combined_path}")

    # -- Step 4: images -------------------------------------------------------
    print("\nStep 4/4: Downloading receipt images into employee folders ...")
    image_mod.download_all(ATTACHMENTS_FILE, run_output_dir)

    print(f"\nAll done. Output: {run_output_dir}")


if __name__ == "__main__":
    main()
