# "C:\Program Files (x86)\Google\Chrome\Application" Chrome Location
driver_path = 'C:\\Program Files (x86)\\chromedriver.exe'

from stat import FILE_ATTRIBUTE_NORMAL
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import json

def writeToJSONFile(data):
    with open('Customer Info.json', 'w') as fp:
        json.dump(data, fp)


# connect to chrome in debugging mode
# chrome.exe --remote-debugging-port=8989 --user-data-dir="C:\Selenium\Chrome_Test_Profile"
opt= Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
driver= webdriver.Chrome(executable_path= driver_path, chrome_options= opt)

# get all windows and switch between windows
print("Parent window title: " + driver.title)
#get current window handle
parent_window = driver.current_window_handle
# get all child windows
chwd = driver.window_handles

customer_info = dict()

# loop through each window performm certain tasks
for child_window in chwd:
    if(child_window!=parent_window):
        driver.switch_to.window(child_window)

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


        customer_info[Order_id.text]= {'Ship_to':Ship_to.text, "Quantity": Quantity.text, "Ship_date": Ship_date.text, "Unit_price": Unit_price.text}


    time.sleep(0.5)


print(customer_info)
writeToJSONFile(customer_info)



