from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from openpyxl import load_workbook

#import helper libraries
import time
import os
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO

webdriver_path = 'chromedriver-win64/chromedriver.exe'
options = Options()
driver = webdriver.Chrome(service=Service(executable_path=webdriver_path), options=options)
driver.set_window_size(1400,1050)

df = pd.read_excel('huawei_project_unique.xlsx')

url = f"https://xlo365.sharepoint.com/:f:/r/sites/ProgramManagement/PMProgramOffice/Shared%20Documents/"

driver.get(url)

time.sleep(120)

for cell_value in df.iloc[:, 0]:
    project_name = str(cell_value).strip()
    print(project_name)
    url = f"https://xlo365.sharepoint.com/:f:/r/sites/ProgramManagement/PMProgramOffice/Shared%20Documents/"
    driver.get(url)
    time.sleep(3)
    search_input = driver.find_element(By.XPATH, '//*[@id="sbcId"]/form/input')
    search_input.send_keys(project_name)
    time.sleep(3)
    search_button = driver.find_element(By.XPATH, '//*[@id="sbcId"]/form/span[6]/button')
    search_button.click()
    time.sleep(3)

    folder_list_cells = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
    file_name_permit_cluster = []
    file_name_permit_feeder = []
    permit_cluster_link = "-"
    permit_feeder_link = "-"


    for count in range(len(folder_list_cells)):
        folder_list_cells_ = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
        folder_cell = folder_list_cells_[count]
        folder_name = folder_cell.get_attribute('textContent')
        if "permit cluster" in folder_name.lower():
            folder_button = folder_cell.find_element(By.TAG_NAME, 'button')
            folder_button.click()   
            time.sleep(3)
            permit_cluster_link = driver.current_url
            file_names = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
            for file_name in file_names:
                file_button = file_name.find_element(By.TAG_NAME, 'button')
                file_name_permit_cluster.append(file_button.get_attribute('textContent'))
            driver.back()
            time.sleep(3)

        if "permit feeder" in folder_name.lower():
            folder_button = folder_cell.find_element(By.TAG_NAME, 'button')
            folder_button.click()   
            time.sleep(3)
            folder_childrens = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
            for count in range(len(folder_childrens)):
                folder_childrens_ = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
                children = folder_childrens_[count]
                child_name = children.get_attribute('textContent')
                if "main" in child_name.lower():
                    children_button = children.find_element(By.TAG_NAME, 'button')
                    children_button.click()
                    time.sleep(3)
                    permit_feeder_link = driver.current_url
                    file_names = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
                    for file_name in file_names:
                        file_button = file_name.find_element(By.TAG_NAME, 'button')
                        file = file_button.get_attribute('textContent')
                        file_name_permit_feeder.append(file)
                    driver.back()
                    time.sleep(3)
                else:
                    permit_feeder_link = driver.current_url
                    file_names = driver.find_elements(By.CLASS_NAME, 'ms-List-cell')
                    for file_name in file_names:
                        file_button = file_name.find_element(By.TAG_NAME, 'button')
                        file = file_button.get_attribute('textContent')
                        file_name_permit_feeder.append(file)    
            driver.back()
            time.sleep(3)
    if len(file_name_permit_feeder) != 0 or len(file_name_permit_cluster) != 0:
        with open("log_sharefolder.csv", "a") as file:
                file.write(f"{project_name}, {file_name_permit_cluster}, {permit_cluster_link}, {file_name_permit_feeder}, {permit_feeder_link}" + "\n")
        print('done')