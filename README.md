# Amazon Invoice Automation 
## From SC (Invoice Data) -> Excel (Generate Invoice) -> SC (Upload Invoice)

## Steps

1. Open Google Chrome in debug mode on localhost:port
2. Open all orders with invoices to be uploaded in seperate tabs
3. Fetch Order_ID and Ship_to details from all opened orders
4. save dict key= Order_id and following values:
    - Ship_to address
    - purchase date
    - Quantity
    - Unit Price
    
5. Write another script to use the details and generate excel invoices
6. Works on Windows Only (win32com libraries)