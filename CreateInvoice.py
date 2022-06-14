from openpyxl import Workbook, load_workbook
import json
import pandas as pd
import pdfkit

EXCEL_FILE = "Base-Template.xlsx"

# with open('Customer Info.json', 'r') as fp:
#   customer_info = json.load(fp)

def create_invoice(Order_id, customer_info):

  invoice_workbook = load_workbook(EXCEL_FILE)
  invoice_worksheet = invoice_workbook.active

  # for customer, details in customer_info.items():
  
  # Invoice number / Order id
  invoice_worksheet['I6'].value = Order_id

  Ship_to = customer_info[Order_id]['Ship_to']
  invoice_worksheet['D6'].value= Ship_to

  Quantity = customer_info[Order_id]['Quantity']
  invoice_worksheet['G15'].value= int(Quantity)
  
  Ship_date = customer_info[Order_id]['Ship_date']
  invoice_worksheet['I7'].value= Ship_date

  Unit_price = customer_info[Order_id]['Unit_price']
  invoice_worksheet['H15'].value= float(Unit_price.replace('£',''))

  xlsx_invoice_path = 'Invoices/'+Order_id+'.xlsx'
  invoice_workbook.save(xlsx_invoice_path)

  # # convert excel to pandas df
  # df = pd.read_excel(xlsx_invoice_path)
  # # convert dataframe to html
  # df.to_html("file.html")
  # # convert html to pdf
  # invoice_path = 'Invoices/'+Order_id+'.pdf'
  # pdfkit.from_file("file.html", invoice_path)
  
  return xlsx_invoice_path


