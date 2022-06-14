from openpyxl import Workbook, load_workbook
import json

EXCEL_FILE = "Notebook-Excel.xlsx"
invoice_workbook = load_workbook(EXCEL_FILE)
invoice_worksheet = invoice_workbook.active

with open('Customer Info.json', 'r') as fp:
  customer_info = json.load(fp)


for customer, details in customer_info.items():

    invoice_worksheet['I6'].value = customer

    Ship_to = details['Ship_to']
    invoice_worksheet['D6'].value= Ship_to

    Quantity = details['Quantity']
    invoice_worksheet['G15'].value= int(Quantity)
    
    Ship_date = details['Ship_date']
    invoice_worksheet['I7'].value= Ship_date

    Unit_price = details['Unit_price']
    invoice_worksheet['H15'].value= float(Unit_price.replace('£',''))

    invoice_workbook.save('Invoices\\'+'Invoice-'+customer+'.xlsx')


