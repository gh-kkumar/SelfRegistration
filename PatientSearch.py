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

FilePath = str(Path().resolve()) + r'\Excel Files\PatientSearch.xlsx'
#FilePath = str(Path().resolve()) + r'\Excel Files\PatientSearchDailyUpdate.xlsx'
Sheet = 'Patient Search Page Data'
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

        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Description')
        Description = RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)

# 2. This is for Patient Search Flow
Sheet = 'Patient Search Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)
Phlebotomist_RowNo = 6

# Patient Search
time.sleep(Seconds)
Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[text() = "Patient Search"]', 60)
Element.click()

for RowIndex1 in range(8, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Run', 6)
    if (RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo) == 'Y'):
        # First Name
        time.sleep(3)
        Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//input[@name = "firstName"]', 60)
        Element.clear()
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'First Name', 6)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
        else:
            Element.clear()
            Element.click()

        # Last Name
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "lastName"]', 60)
        Element.clear()
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Last Name', 6)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
        else:
            Element.clear()
            Element.click()

        # MRN
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "patientId"]', 60)
        Element.clear()
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'MRN', 6)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
        else:
            Element.clear()
            Element.click()

        # DOB
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "dateOfBirth"]', 60)
        Element.clear()
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'DOB', 6)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
        else:
            Element.clear()
            Element.click()

        # Phone
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "phone"]', 60)
        Element.clear()
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Phone', 6)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 6)))
        else:
            Element.clear()
            Element.click()

        # Search Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Search"]', 60)
        Element.click()

        time.sleep(Seconds)
        # Table
        Element = '//div[2]/div/div/table'
        TableRows = Element + '/tbody/tr'
        row_count = len(driver.find_elements_by_xpath(TableRows))
        CellVal = ['', '', '', '']
        if (row_count > 0):
            # Click Triangle
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-button-menu//lightning-primitive-icon', 60)
            Element.click()

            # Create Order Link
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Create Order"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Next Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Next Button
            time.sleep(3)
            Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            OM.Place_Order(driver, FilePath, Sheet, RowIndex1, RowCount, Seconds, Seconds1, Phlebotomist_RowNo, 6, 'Patient Search')
            Phlebotomist_RowNo += 1

time.sleep(2)
driver.quit()
