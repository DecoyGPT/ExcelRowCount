import openpyxl
import sys

def count_rows_with_data(workbook_path):
    # Load the workbook
    wb = openpyxl.load_workbook(workbook_path)

    grand_total_rows = 0

    # Iterate over all sheets in the workbook
    for sheet in wb.worksheets:
        sheet_total_rows = 0
        # Iterate over all rows in the sheet
        for row in sheet.iter_rows():
            # Check if the row has data (non-empty cells)
            if any(cell.value for cell in row):
                sheet_total_rows += 1

        print(f"{sheet.title} = {sheet_total_rows} rows")
        grand_total_rows += sheet_total_rows

    print("\nTotal Rows:", grand_total_rows, "rows")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_name.py <path_to_workbook>")
        sys.exit(1)

    workbook_path = sys.argv[1]
    count_rows_with_data(workbook_path)