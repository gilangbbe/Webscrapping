import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service

#import helper libraries
import time
import os
from io import StringIO
import util
import pyautogui
import pyperclip

webdriver_path = 'chromedriver-win64/chromedriver.exe'
options = Options()
options.add_argument("--user-data-dir=C:/Users/777707240052/AppData/Local/Google/Chrome/User Data")
options.add_argument("--start-maximized") 
driver = webdriver.Chrome(service=Service(executable_path=webdriver_path), options=options)

excel_file_path = 'EvidenceSiteAccess.xlsx'
df = pd.read_excel(excel_file_path, sheet_name='Sheet1')

for i, value in df.iterrows():
    driver.get(f'https://drive.google.com/drive/search?q=type:pdf%20parent:18LyekC8cc8pvIIQA6-u38bj4cNifw1WY%20title:{value['SiteID']}')
    time.sleep(5)

    pyautogui.click(563,502)
    time.sleep(2)
    pyautogui.click(864,360)
    time.sleep(1)
    link = pyperclip.paste()

    with open("link.csv", "a") as file:
        file.write(f"{value['SiteID']}, {link}" + "\n")