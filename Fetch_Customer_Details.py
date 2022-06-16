# "C:\Program Files (x86)\Google\Chrome\Application" Chrome Location
# launch chrome in debugging mode on port 8989
# chrome.exe --remote-debugging-port=8989 --user-data-dir="C:\Selenium\Chrome_Test_Profile" (WINDOWS)
# google-chrome --remote-debugging-port=8989 --user-data-dir="/home/jarawr/google_chrome_test" (UBUNTU)

LINUX_CHROMEDRIVER = '/home/jarawr/chromedriver'
WIN_CHROMEDRIVER = 'C:\Program Files (x86)\chromedriver.exe'

from distutils.command.upload import upload
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json
from collections import OrderedDict
from CreateInvoice import create_invoice
from selenium.common.exceptions import NoSuchElementException        

def check_exists_by_xpath(chromedriver, xpath):
    try:
        chromedriver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

opt= Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
driver= webdriver.Chrome(executable_path= WIN_CHROMEDRIVER, chrome_options= opt)

# get all windows and switch between windows

# get current window handle
driver.switch_to.window(driver.current_window_handle)
# get all child windows
chwd = driver.window_handles

customer_info = OrderedDict()
i = 0
# loop through each window performm certain tasks
for child_window in chwd:
    i+=1
    # Benjamin Johns
    # 24
    # AUSTEN DRIVE
    # WESTON-SUPER-MARE
    # BS22 7UY
    # United Kingdom
    Ship_to = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div")
    Ship_to = Ship_to.text
    # print(Ship_to.text)

    # 026-0968862-2197164
    Order_id = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]")
    Order_id = Order_id.text
    # print(Order_id.text)

    # 2
    Quantity = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]")
    Quantity = Quantity.text
    # print(Quantity.text)

    # Wed, 1 Jun 2022
    Ship_date = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[2]/div[1]/div/div/div/div/div/div[1]/table/tbody/tr[1]/td[2]/span")
    Ship_date = Ship_date.text
    # print(Ship_date.text)

    # £10.95
    Unit_price = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span")
    Unit_price = Unit_price.text
    # print(Unit_price.text)

    # Shipping Price
    """Pending Start"""
    # if driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span"):
    #     # take value
    #     Ship_price = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span").text
    # else:
    #     Ship_price = '0' 
    """Pending End"""

    # keep saving all order id's with relevant stats to a dict of order id's ---> dict = {'Order_id_1':{'a':1,'b':2,'c':3}, 'Order_id_2':{...}...}
    customer_info[Order_id]= {'Ship_to':Ship_to, "Quantity": Quantity, "Ship_date": Ship_date, "Unit_price": Unit_price}

    # Create and save invoice
    pdf_invoice_path = create_invoice(Order_id, customer_info)


    while not check_exists_by_xpath(driver, "/html/body/div[3]/div/div/div/div/div[2]/form/span[2]"):
        # Steps: Refresh -> Upload -> Browse
            # if document uploaded then loop breaks
            # else keep repeating the loop
        print(Order_id)
        # refresh the page before uploading any data
        # driver.refresh()
        # wait for a few seconds to make sure page loads
        # time.sleep(3)

        # Wait for upload button to appear and click it
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[1]/div[2]/div[2]/span[1]/span'))
            ).click()

        # print('CLICKED THE BUTTON TO OPEN PROMPT')
        
        # Wait for browse button to appear, then send the generated invoice via sendkeys method
        browse_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "invoice-file"))
            ).send_keys(pdf_invoice_path)

        # # fill in the document number box
        # document_number = driver.find_element(By.XPATH, '//*[@id="invoice_number_input"]').send_keys(Order_id)
        
        # # wait for upload to process and document number to appear
        # time.sleep(1)

    print('Invoice Number'+str(i)+' Generated: ', pdf_invoice_path)

    # break
    driver.switch_to.window(child_window)

