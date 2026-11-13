import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter


# === Configuration ===
excel_file = r"C:\Users\Sudhanshu.t\Downloads\Expense Report June.xlsx"

# sheet_name = 'Dates Append'
base_output_dir = r"C:\Users\Sudhanshu.t\Downloads\Expense June 26_output"
verticals = ['FS-CH', 'NUT', 'FS-KA', 'DIR']
second_filter_column = 'Region'  # 🔁 Replace this with your actual second filter column


def load_input_data(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == '.csv':
        return pd.read_csv(file_path)
    if extension in {'.xlsx', '.xls', '.xlsm'}:
        return pd.read_excel(file_path)

    raise ValueError(f"Unsupported file format: {extension}")


def process_expenses(file_path, output_dir, verticals_to_process, second_filter, df=None):
    if df is None:
        df = load_input_data(file_path)

    if 'Key' in df.columns:
        df = df.drop(columns=['Key'])

    for vertical in verticals_to_process:
        vertical_df = df[df['BSV'] == vertical]

        if vertical_df.empty:
            continue

        vertical_folder = os.path.join(output_dir, vertical.lower().replace('-', '').replace(' ', ''))
        os.makedirs(vertical_folder, exist_ok=True)

        second_level_values = vertical_df[second_filter].dropna().unique()

        for value in second_level_values:
            second_level_df = vertical_df[vertical_df[second_filter] == value]
            second_level_folder = os.path.join(vertical_folder, str(value).strip().replace(' ', '_'))
            os.makedirs(second_level_folder, exist_ok=True)

            employees = second_level_df['employee name'].dropna().unique()

            for employee in employees:
                employee_df = second_level_df[second_level_df['employee name'] == employee]
                employee_name_clean = str(employee).replace(' ', '_')
                employee_file_path = os.path.join(second_level_folder, f"{employee_name_clean}.xlsx")

                employee_df.to_excel(employee_file_path, index=False)

                wb = load_workbook(employee_file_path)
                ws = wb.active

                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

                for row in ws.iter_rows():
                    for cell in row:
                        cell.border = thin_border

                for col in ws.columns:
                    max_length = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except Exception:
                            pass
                    adjusted_width = max_length + 2
                    ws.column_dimensions[col_letter].width = adjusted_width

                wb.save(employee_file_path)


def main():
    process_expenses(excel_file, base_output_dir, verticals, second_filter_column)


if __name__ == "__main__":
    main()
