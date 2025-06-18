import csv
import openpyxl as xlsx

def get_csv_headers(path):
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        return next(reader)

def get_xlsx_headers(path):
    wb = xlsx.load_workbook(path, read_only=True)
    sheet = wb.active
    return [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
