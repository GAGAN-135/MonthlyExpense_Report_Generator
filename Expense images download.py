import os
import re
import uuid
import pandas as pd
import requests

# =============================== Configuration ===============================
# File containing the CDN links (the attachments export). Supports .csv / .xlsx
attachments_file = r"C:\Users\Sudhanshu.t\Downloads\Expense Report June attachments (2).csv"

# The "Split Expenses.py" output folder. Each employee has a "<Employee_Name>.xlsx"
# saved inside its region folder. Images are dropped into a folder named after the
# employee, right next to that xlsx, i.e.
#   <split_output_dir>/<vertical>/<region>/<Employee_Name>/<image>
split_output_dir = r"C:\Users\Sudhanshu.t\Downloads\Expense June 26_output"

# Column names in the attachments file
link_column = "file_link"          # Column with the image URL
name_column = "customer_name"      # Column with the employee/customer name (== "employee name")

# Attachments whose name doesn't match any split-output employee are still saved
# here so nothing is lost.
unmatched_folder_name = "_Unmatched_Attachments"
# =============================================================================


def normalize_name(name):
    """Normalize a name for matching: underscores->spaces, collapse spaces, lowercase."""
    return re.sub(r"\s+", " ", str(name).replace("_", " ")).strip().lower()


def build_employee_folder_map(output_dir):
    """Scan the Split Expenses output and map each employee to the folder that
    holds their expense file. Returns {normalized_name: (region_folder, stem)}."""
    emp_map = {}
    for root, _dirs, files in os.walk(output_dir):
        for f in files:
            if f.lower().endswith((".xlsx", ".xls", ".xlsm")):
                stem = os.path.splitext(f)[0]
                emp_map[normalize_name(stem)] = (root, stem)
    return emp_map


def load_attachments(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension == ".csv":
        return pd.read_csv(file_path)
    if extension in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(file_path)
    raise ValueError(f"Unsupported attachments file format: {extension}")


def resolve_target_folder(name, emp_map):
    """Return the folder where this employee's images should be saved.
    Matched -> <region_folder>/<Employee_Name>/. Unmatched -> _Unmatched/<name>/."""
    key = normalize_name(name)
    if key in emp_map:
        region_folder, stem = emp_map[key]
        return os.path.join(region_folder, stem)
    safe_name = str(name).strip().replace(" ", "_") or "Unknown"
    return os.path.join(split_output_dir, unmatched_folder_name, safe_name)


def download_image(url, target_folder, image_name):
    try:
        response = requests.get(url, stream=True, timeout=60)
        if response.status_code == 200:
            os.makedirs(target_folder, exist_ok=True)

            # Keep the original extension from the URL when available
            file_root, file_ext = os.path.splitext(image_name)
            unique_image_name = f"{file_root}_{uuid.uuid4().hex[:6]}{file_ext}"

            image_path = os.path.join(target_folder, unique_image_name)
            with open(image_path, "wb") as file:
                file.write(response.content)
            print(f"Downloaded {unique_image_name} -> {target_folder}")
            return True
        print(f"Failed to download {url} (status code: {response.status_code})")
    except Exception as e:
        print(f"Error occurred while downloading {url}: {e}")
    return False


def download_all(attach_file, output_dir, link_col=link_column, name_col=name_column):
    """Download every attachment into its employee's folder inside the split output.
    Returns (downloaded, failed, unmatched_names). Reusable by the orchestrator."""
    if not os.path.isdir(output_dir):
        print(f"Error: Split output folder not found: '{output_dir}'")
        print("Run the split step first so the employee folders exist.")
        return 0, 0, set()

    try:
        df = load_attachments(attach_file)
    except FileNotFoundError:
        print(f"Error: The file '{attach_file}' was not found.")
        return 0, 0, set()

    if link_col not in df.columns or name_col not in df.columns:
        print(f"Error: Required columns '{link_col}'/'{name_col}' not found.")
        print(f"Available columns: {list(df.columns)}")
        return 0, 0, set()

    emp_map = build_employee_folder_map(output_dir)
    print(f"Found {len(emp_map)} employee folders in the split output.\n")

    downloaded, failed, unmatched_names = 0, 0, set()

    for _index, row in df.iterrows():
        file_link = row[link_col]
        employee_name = row[name_col]

        if pd.isna(file_link) or str(file_link).strip() == "":
            continue

        if normalize_name(employee_name) not in emp_map:
            unmatched_names.add(str(employee_name))

        target_folder = resolve_target_folder(employee_name, emp_map)

        # Build the base image name from date + amount, keeping the URL's extension
        expense_date = pd.to_datetime(row["expense_date"]).strftime("%d-%m-%Y")
        expense = str(row["expense"])
        url_ext = os.path.splitext(str(file_link).split("?")[0])[1] or ".jpg"
        image_name = f"{expense_date} {expense}{url_ext}"

        if download_image(file_link, target_folder, image_name):
            downloaded += 1
        else:
            failed += 1

    print(f"\nDone. Downloaded: {downloaded}, Failed: {failed}")
    if unmatched_names:
        print(f"\n{len(unmatched_names)} name(s) had no matching employee folder "
              f"(saved under '{unmatched_folder_name}'):")
        for name in sorted(unmatched_names):
            print(f"  - {name}")
    return downloaded, failed, unmatched_names


def main():
    download_all(attachments_file, split_output_dir)


if __name__ == "__main__":
    main()
