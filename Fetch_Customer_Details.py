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
import os
import json
from collections import OrderedDict
from CreateInvoice import create_invoice


def writeToJSONFile(data):
    with open('Customer Info.json', 'w') as fp:
        json.dump(data, fp)



opt= Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
driver= webdriver.Chrome(executable_path= WIN_CHROMEDRIVER, chrome_options= opt)

# get all windows and switch between windows

# get current window handle
driver.switch_to.window(driver.current_window_handle)
# get all child windows
chwd = driver.window_handles

customer_info = OrderedDict()

# loop through each window performm certain tasks
for child_window in chwd:

    # Benjamin Johns
    # 24
    # AUSTEN DRIVE
    # WESTON-SUPER-MARE
    # BS22 7UY
    # United Kingdom
    Ship_to = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[1]/table/tbody/tr[1]/td/span/span[3]/div/div")
    # print(Ship_to.text)

    # 026-0968862-2197164
    Order_id = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[1]/div[1]/div/span[5]")
    # print(Order_id.text)

    # 2
    Quantity = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[5]")
    # print(Quantity.text)

    # Wed, 1 Jun 2022
    Ship_date = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[2]/div[1]/div/div/div/div/div/div[1]/table/tbody/tr[1]/td[2]/span")
    # print(Ship_date.text)

    # £10.95
    Unit_price = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[6]/div/table/tbody/tr/td[6]/span")
    # print(Unit_price.text)

    # keep saving all order id's with relevant stats to a dict of order id's ---> dict = {'Order_id_1':{'a':1,'b':2,'c':3}, 'Order_id_2':{...}...}
    customer_info[Order_id.text]= {'Ship_to':Ship_to.text, "Quantity": Quantity.text, "Ship_date": Ship_date.text, "Unit_price": Unit_price.text}

    # Create and save invoice
    pdf_invoice_path = create_invoice(Order_id.text, customer_info)

    # Click on Upload button and wait
    # upload_button = driver.find_element(By.XPATH, "//*[@id='MYO-app']/div/div[1]/div[1]/div/div[1]/div[2]/div[2]/span[1]")
    # upload_button.click()
    
    # Wait for browse button to appear, then send the generated invoice via sendkeys method
    # browse_button = WebDriverWait(driver, 5).until(
    #     EC.presence_of_element_located((By.NAME, "invoice-file"))
    #     ).send_keys(pdf_invoice_path)

    print('Path of upload pdf: ', pdf_invoice_path)

    # break
    driver.switch_to.window(child_window)

