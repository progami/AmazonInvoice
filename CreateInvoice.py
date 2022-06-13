from openpyxl import Workbook, load_workbook
import json

EXCEL_FILE = "Notebook-Excel.xlsx"

with open('Customer Info.json', 'r') as fp:
  customer_info = json.load(fp)

# invoice_workbook = load_workbook(EXCEL_FILE)
# invoice_worksheet = invoice_workbook.active

# invoice_worksheet['D6'].value = "Test Address"

# invoice_workbook.save(EXCEL_FILE)


