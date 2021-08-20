import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import OrderManagement as OM
import os
#import win32com.client as comclt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path
from selenium.webdriver.common import keys

FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Url')
Url = str(RWDE.ReadData(FilePath, Sheet, 3, ColumnNo))

driver = webdriver.Chrome(executable_path = str(Path().resolve()) + r'\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)

#1. This is for HCP Login Page

FilePath = str(Path().resolve()) + r'\Excel Files\SearchOrder.xlsx'
#FilePath = str(Path().resolve()) + r'\Excel Files\PatientSearchDailyUpdate.xlsx'
Sheet = 'SearchOrder Page Data'
RowCount = 3
Seconds = 1
Seconds1 = 300 / 1000

for RowIndex in range(2, RowCount + 1):
    Column = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
    if (RWDE.ReadData(FilePath, Sheet, RowIndex, Column) == 'Y'):
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Name')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Password')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//span[text() = "Log in"]', 60)
        Element.click()