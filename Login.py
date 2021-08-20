import time
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import WebElementReusability as WER
import openpyxl
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path

def HCPOrClinicalStaff_Login():
    FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
    Sheet = 'Portal Urls'
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Url')
    Url = str(RWDE.ReadData(FilePath, Sheet, 3, ColumnNo))

    driver = webdriver.Chrome(executable_path=str(Path().resolve()) + r'\Browser\chromedriver_win32\chromedriver')
    driver.maximize_window()
    driver.get(Url)

    # 1. This is for HCP Login Page

    FilePath = str(Path().resolve()) + r'\Excel Files\HCPLogin.xlsx'
    Sheet = 'Login Page Data'
    RowCount = RWDE.RowCount(FilePath, Sheet)
    Seconds = 1  # 300 / 1000

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
            Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//span[. = "Log in"]', 60)
            Element.click()

            time.sleep(7)
            print(driver.title)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Expected Result')
            ColumnNo1 = RWDE.FindColumnNoByName(FilePath, Sheet, 'Actual Result')
            ColumnNo2 = RWDE.FindColumnNoByName(FilePath, Sheet, 'Result')
            if (driver.title == 'Login'):
                Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div/div/span/div/div', 60)
                if (RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == Element.text):
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo1, Element.text)
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo2, 'Pass')
                else:
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo1, Element.text)
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo2, 'Fail')

                driver.execute_script('arguments[0].innerHTML = ""', Element)
                Element = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
                Element.clear()

                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
                Element.clear()
            elif (driver.title == 'Guardant Health'):
                if (RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == driver.title):
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo1, driver.title)
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ColumnNo2, 'Pass')

