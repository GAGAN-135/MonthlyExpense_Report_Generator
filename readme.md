# Monthly Expense Report Generator

A Python-based tool for processing, organizing, and managing monthly employee expense reports. This project automates downloading expense receipt images and splitting expense data by organizational hierarchy (verticals, regions, and employees).

## 📋 Overview

This tool handles four main tasks, wired together by a one-click runner (`run_pipeline.py`):
1. **Enrich** the monthly expense sheet with `BSV` + `Region`, looked up from the online
   employee master (**Channelier Employees.xlsx**, synced locally via OneDrive)
2. **Split and organize** expense data into individual formatted Excel reports by employee
3. **Combined overview** — one summary sheet (name, id, BSV, region, total) across all employees
4. **Download expense receipts** from URLs into each employee's folder

The two underlying scripts (`Split Expenses.py`, `Expense images download.py`) still work
standalone; `run_pipeline.py` just runs all three steps in order.

## 🚀 Features

- **Automated Image Download**: Downloads expense receipt images from URLs with automatic folder organization by employee
- **Hierarchical Data Splitting**: Organizes expense data by vertical (business segment) → region → employee
- **Formatted Excel Export**: Creates professionally formatted Excel files with borders and auto-sized columns
- **Duplicate Prevention**: Uses unique identifiers to prevent file overwrites
- **Batch Processing**: Processes multiple employees and regions in a single run

## 📦 Requirements

- Python 3.7+
- pandas
- requests
- openpyxl

Install dependencies using:
```bash
pip install -r requirements.txt
```

## ⚙️ Setup

### 1. Create Virtual Environment

Run the setup script to create and activate a Python virtual environment:

**On Windows (PowerShell):**
```powershell
.\setup_venv.ps1
```

This will:
- Create a `venv` folder
- Install all required packages from `requirements.txt`
- Activate the virtual environment

### 2. Configure the Scripts

Before running, you need to update file paths in both scripts:

#### For `Expense images download.py`:
```python
excel_file = r"C:\Path\To\Your\Expenses_File.xlsx"
save_parent_folder = r"C:\Path\To\Save\Downloaded\Images"
column_name = "file_link"           # Column containing image URLs
employee_column = "customer_name"   # Column containing employee names
```

#### For `Split Expenses.py`:
```python
excel_file = r"C:\Path\To\Your\Expenses_Apr'26.xlsx"
base_output_dir = r"C:\Path\To\Output\Directory"
verticals = ['FS-CH', 'NUT', 'FS-KA', 'DIR']  # Your business segments
second_filter_column = 'Region'  # Secondary grouping column
```

## 🚀 Quick Start (one-click — recommended)

Run the whole pipeline with a single command:

```powershell
python run_pipeline.py
```

It performs **enrich → split → combined overview → download** in order and prints a per-step
summary (including any employees missing from the master).

### One-time setup: the employee master (OneDrive)

`run_pipeline.py` reads BSV + Region from **Channelier Employees.xlsx**, which lives on the
analytics account's OneDrive. You need a real local copy (a `.url` web-shortcut will NOT work):

1. Have the file's owner **share the folder** that contains it with you.
2. In OneDrive (browser) → **Shared** → open that **folder** → **Add shortcut to My files**
   (a folder shortcut syncs the real file; a single-file shortcut only makes a `.url` link).
3. Right-click the synced file → **Always keep on this device**.
4. Point `EMPLOYEE_MASTER_FILE` at that local path.

Because it's the synced copy, edits to the online master (e.g. adding a new employee) are
picked up automatically on the next run.

### Configure `run_pipeline.py`

Edit only the CONFIG block at the top:

```python
EXPENSE_FILE         = r"C:\...\Expense Report June.xlsx"                 # missing BSV/Region
ATTACHMENTS_FILE     = r"C:\...\Expense Report June attachments.csv"      # image CDN links
EMPLOYEE_MASTER_FILE = r"C:\Users\<you>\OneDrive - <Org>\Channelier Employees.xlsx"
OUTPUT_DIR           = r"C:\...\Expense June 26_output"                   # timestamp appended per run

# Lookup mapping (matched to Channelier Employees.xlsx)
EXPENSE_KEY_COL = "employee code"          # join key in the expense sheet (EM366)
MASTER_KEY_COL  = "Employee ID"            # join key in the master        (EM366)
MASTER_BSV_COL  = "Business Sub Vertical"  # FS-CH / FS-KA / NUT / DIR  (NOT "Business Vertical")
MASTER_REGION_COL = "Region"               # North-1, West, East, ...
DEFAULT_REGION  = "0"                       # bucket for valid-BSV but blank-Region employees

# Combined overview sheet
EMP_NAME_COL        = "employee name"      # employee-name column in the expense sheet
TOTAL_AMOUNT_COL    = "total amount"       # amount column summed per employee
COMBINED_FOLDER_NAME = "Combined Expense"  # subfolder that holds the summary sheet
```

### Output: timestamped folders

Each run writes to its own date-time-stamped folder, so images never accumulate or duplicate
across runs:

```
Expense July 26_output_2026-08-14_1330/
├── Combined Expense/
│   └── Combined Expense Summary.xlsx   # one row per employee: name, id, BSV, region, total
├── fsch/North-3/Anish_Kumar.xlsx
├── fsch/North-3/Anish_Kumar/           # this employee's receipt images
│   ├── 12-07-2026 100.0_a1b2c3.jpeg
│   └── ...
└── ...
```

The **Combined Expense Summary** is sorted by BSV → region → total (highest first) and ends
with a bold grand-total row — a quick overview of spend per employee.

Employees missing from the master are listed in the console and skipped; unmatched
attachments (if any) land in an `_Unmatched_Attachments/` folder so nothing is lost.

---

## 📖 Usage (standalone scripts)

> The steps below run the two scripts individually. Most users can skip this and use
> `run_pipeline.py` above. Note: run standalone, `Split Expenses.py` expects the expense sheet
> to already contain `BSV` and `Region` columns (which the one-click runner adds for you).

### Step 1: Download Expense Receipt Images

Run the expense images download script:

```powershell
python "Expense images download.py"
```

**What it does:**
- Reads an Excel file containing expense data with image URLs
- Downloads each image from the provided URL
- Organizes files into folders by employee name
- Names files as: `DD-MM-YYYY_ExpenseAmount.jpg` (with unique suffix to prevent overwrites)
- Prints download progress and any errors encountered

**Output structure:**
```
Saved_Images_Folder/
├── Employee_Name_1/
│   ├── 15-04-2026_5000_a1b2c3.jpg
│   └── 20-04-2026_3500_d4e5f6.jpg
└── Employee_Name_2/
    └── 10-04-2026_2500_g7h8i9.jpg
```

### Step 2: Split and Format Expense Data

Run the split expenses script:

```powershell
python "Split Expenses.py"
```

**What it does:**
- Reads the main expense data file
- Filters and splits data by vertical (business segment)
- Further organizes by region
- Creates individual Excel files for each employee
- Applies formatting: borders around all cells, auto-sized columns
- Removes any "Key" column if present
- Saves formatted files in hierarchical folder structure

**Output structure:**
```
Base_Output_Dir/
├── fsch/
│   ├── Region_A/
│   │   ├── Employee_1.xlsx
│   │   └── Employee_2.xlsx
│   └── Region_B/
│       └── Employee_3.xlsx
├── nut/
│   └── ...
├── fska/
│   └── ...
└── dir/
    └── ...
```

## 🔄 Workflow

```
┌──────────────────┐     ┌────────────────────────────┐
│  Expense sheet   │     │  Channelier Employees.xlsx │
│ (no BSV/Region)  │     │  (OneDrive-synced master)  │
└────────┬─────────┘     └─────────────┬──────────────┘
         │                             │
         └───────────┬─────────────────┘
                     ▼
          ┌────────────────────────┐
          │ 1. Enrich: lookup BSV  │
          │    + Region by emp code│
          └───────────┬────────────┘
                      ▼
          ┌────────────────────────┐     ┌──────────────────┐
          │ 2. Split by Vertical   │     │  Attachments CSV │
          │    + Region + Employee │     │  (image links)   │
          └───────────┬────────────┘     └────────┬─────────┘
                      ▼                            │
          ┌────────────────────────┐              │
          │ 3. Combined overview   │              │
          │    (summary sheet)     │              │
          └───────────┬────────────┘              │
                      ▼                            │
          ┌────────────────────────┐              │
          │  Employee .xlsx files  │◄─────────────┘
          │  + receipt images   4. Download into each folder
          └────────────────────────┘
   (all inside one timestamped output folder per run)
```

### Complete Workflow Steps:

1. **Refresh the master**: nothing to do — the OneDrive-synced `Channelier Employees.xlsx`
   is always current. Just make sure new employees have `Employee ID`, `Business Sub Vertical`,
   and `Region` filled in.

2. **Update configuration**: set the four paths + column mapping in `run_pipeline.py` (once).

3. **Run**: `python run_pipeline.py` — enrich, split, combined overview, and download happen in order.

4. **Review output**:
   - Console prints rows enriched, employees summarised, images downloaded, and any employees missing from the master.
   - Open the new `..._output_<timestamp>` folder: a `Combined Expense` summary, plus each employee's `.xlsx` and receipts folder.

## 🛠️ Customization

### Change File Naming Format
In `Expense images download.py`, modify the `image_name` construction:
```python
image_name = f"{expense_date} {expense}.jpg"
```

### Change Organizational Hierarchy
In `Split Expenses.py`, modify the `verticals` list and `second_filter_column`:
```python
verticals = ['YOUR', 'BUSINESS', 'SEGMENTS']
second_filter_column = 'Your_Grouping_Column'
```

### Change Output Location
Update `base_output_dir` to save to a different location

## ⚠️ Important Notes

- **File Paths**: Use raw strings (r"...") for Windows paths with backslashes
- **Column Names**: Ensure Excel column names exactly match those specified in the scripts
- **Network**: Image download requires internet connection and accessible URLs
- **Permissions**: Ensure write permissions in output directories
- **Large Files**: For large datasets, execution may take several minutes

## 🐛 Troubleshooting

### Script fails with "FileNotFoundError"
- Verify the Excel file path exists and is correct
- Check that column names match exactly

### Images not downloading
- Verify URLs are valid and accessible
- Check internet connection
- Review error messages in console output

### Permission denied errors
- Ensure output directories have write permissions
- Close any open Excel files in the output location
- Run PowerShell as Administrator if needed

### Empty output folders
- Check that filter values (verticals, regions) exist in your data
- Verify column names match the actual Excel file
- Review console output for filtering details

### Employee master reads as a `.url` / "not the actual spreadsheet"
- "Add shortcut to My files" on a **single file** only creates a `.url` web link, not the file.
  Sync the **folder** that contains it instead (see Quick Start → One-time setup), then point
  `EMPLOYEE_MASTER_FILE` at the real synced `.xlsx`.

### An employee is "missing from the master" / skipped
- They aren't in `Channelier Employees.xlsx`, or their `Employee ID` doesn't match the expense
  sheet's `employee code`. Add/fix them in the master (with `Business Sub Vertical` + `Region`)
  and re-run.

## 📝 File Structure

```
MonthlyExpense_Report_Generator/
├── run_pipeline.py               # ⭐ One-click: enrich → split → download
├── Split Expenses.py             # Splits and formats expense data (also importable)
├── Expense images download.py    # Downloads receipt images (also importable)
├── requirements.txt              # Python dependencies
├── setup_venv.ps1                # Virtual environment setup script
└── readme.md                      # This file
```

## 🔗 Requirements Versions

- pandas: Latest stable
- requests: Latest stable
- openpyxl: Latest stable

## 📞 Support

For issues with:
- **File paths**: Verify they use raw strings and exist on your system
- **Excel format**: Ensure the file is `.xlsx` (not `.xls`)
- **Column names**: Print `df.columns` to see exact column names
- **Dependencies**: Run `pip install -r requirements.txt` again

---

**Last Updated**: July 2026
