from openpyxl import Workbook, load_workbook
import win32com.client
import os


def excel_to_pdf(xlsx_invoice_path):
  try:
    xl_file = win32com.client.Dispatch("Excel.Application")
    xl_file.Visible = False

    wb = xl_file.Workbooks.Open(xlsx_invoice_path)

    path_to_pdf = xlsx_invoice_path.replace('.xlsx', '.pdf')

    # print('Path to PDF: ', path_to_pdf)

    print_area = 'A1:I22'
    
    ws = wb.Worksheets[0]
    
    ws.PageSetup.Zoom = False
    
    ws.PageSetup.FitToPagesTall = 1

    ws.PageSetup.FitToPagesWide = 1

    ws.PageSetup.PrintArea = print_area
    
    wb.WorkSheets([1]).Select()
    
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

    wb.Close(True)

    return path_to_pdf
    
  except Exception as e:
    print('error occoured trying to convert to pdf: ', e)
    wb.Close(True)

# with open('Customer Info.json', 'r') as fp:
#   customer_info = json.load(fp)

def create_invoice(Order_id, customer_info):
  
  TEMPLATE_FILE = "Base-Template.xlsx"
  # Open template invoice
  invoice_workbook = load_workbook(TEMPLATE_FILE)
  invoice_worksheet = invoice_workbook.active
  
  # Change data in template invoice according to what you got from customer_info
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

  xlsx_invoice_path = os.getcwd()+'\\'+'Invoices\\'+Order_id+'.xlsx'

  # Save the modified invoice as an excel file
  invoice_workbook.save(xlsx_invoice_path)
  # close the workbook
  invoice_workbook.close()
  # convert the excel file into a pdf file, pass in the path of the excel file
  pdf_invoice_path= excel_to_pdf(xlsx_invoice_path)
  # return the path of the converted pdf file
  return pdf_invoice_path


