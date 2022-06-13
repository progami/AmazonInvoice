# "C:\Program Files (x86)\Google\Chrome\Application" Chrome Location
driver_path = 'C:\\Program Files (x86)\\chromedriver.exe'

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time

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

# loop through each window performm certain tasks
for child_window in chwd:
    if(child_window!=parent_window):
        driver.switch_to.window(child_window)
        print("Child window title: " + driver.title)
    time.sleep(1)




