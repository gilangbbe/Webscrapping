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
import pandas as pd

webdriver_path = 'chromedriver-win64/chromedriver.exe'
options = Options()
options.add_argument("--user-data-dir=C:/Users/777707240052/AppData/Local/Google/Chrome/User Data")
options.add_argument("--start-maximized") 
driver = webdriver.Chrome(service=Service(executable_path=webdriver_path), options=options)

df = pd.read_excel('EvidenceSiteAccess.xlsx')

def get_scroll_dimension(axis):
    return driver.execute_script(f"return document.body.parentNode.scroll{axis}")

failed_site_id = [190694109,
    190693109,
    190692109,
    190687109,
    190686109,
    190684109,
    190683109]


# for cell_value in df.iloc[:, 0]:
for cell_value in failed_site_id:
    site_id = str(cell_value).strip()
    driver.get(f"https://www.appsheet.com/start/82a17c9b-6848-41ad-96cc-45dc5c6256dc#appName=SurveyValidasiFOTSEL51-6021383&group=%5B%5D&row={site_id}&sort=%5B%5D&table=Sheet1&view=Sheet1_Detail")
    print(site_id)
    time.sleep(6)
    site_id = driver.find_element(By.XPATH, '//*[@id="scroller"]/div/div/section/div[4]/div[2]/div[1]/div/span').get_attribute('textContent')
    # get the page scroll dimensions
    # width = get_scroll_dimension("Width")
    # height = get_scroll_dimension("Height")

    # util.fullpage_screenshot(driver, f"{site_id}.png")
    # driver.save_screenshot(f'{site_id}.png')
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write(f'{site_id}')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)

    # next_button = driver.find_element(By.XPATH, '//*[@id="RowDetailPage468bb7c8"]/div/div/div/div[1]/div/div/div[2]/div/span[2]/div/i')
    