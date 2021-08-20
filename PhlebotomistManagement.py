import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
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

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-notifications')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver', options=chrome_options)
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\PhlebotomistManagement.xlsx'
#FilePath = str(Path().resolve()) + '\Excel Files\PhlebotomistManagementDailyUpdate.xlsx'
Seconds = 1

#1. This is for SFDC Login

Sheet = 'Phlebotomist page Data'
RowCount = 2
Seconds = 1

for RowIndex in range(2, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
    if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Name')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Password')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[contains(text(),"Log in")]', 60)
        Element.click()

Sheet1 = 'Phlebotomist page Data'
RowCount1 = RWDE.RowCount(FilePath, Sheet1)
for Rowindex1 in range(6, RowCount1 + 1):
    RowNo = 5
    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Run', RowNo)
    if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
        # Order Number
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "ordernumber"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Order Number', RowNo)
        if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo))
        else:
            Element.send_keys('')

        # Scan Order Button
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Scan Order"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Collect Specimen Samples', RowNo)
        if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo)) == 'Y'):
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "bckId"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'BCK ID', RowNo)
            if (str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo))
                BCKId = RWDE.ReadData(FilePath, Sheet1, Rowindex1, ColumnNo)
            else:
                Element.send_keys('')

            # Next Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[text() = "Next"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Error Message
            time.sleep(Seconds)
            ErrorExists = WER.check_exists_by_xpath(driver, '/html/body/div[4]/div/div/div/div/div/span')
            # ErrorExists1 = WER.check_exists_by_xpath(driver, '//div[6]/div/div/div/div//span')
            if (ErrorExists == True):
                time.sleep(1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "bckId"]', 60)
                Element.clear()
            else:
                # Tube1 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb1tube"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(false);', Element)
                Tube1 = Element.get_attribute('value')

                # Tube2 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb2tube"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(false);', Element)
                Tube2 = Element.get_attribute('value')

                # Tube3 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb3tube"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(false);', Element)
                Tube3 = Element.get_attribute('value')

                # Tube4 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb4tube"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(false);', Element)
                Tube4 = Element.get_attribute('value')

                # Submit Button
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[text() = "Submit"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(false);', Element)
                driver.execute_script('arguments[0].click();', Element)

                # Result Message
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/div/div/div/div/div/span', 60)
                ResultMessage = Element.text

                FilePath1 = str(Path().resolve()) + r'\Excel Files\BSMManagement.xlsx'
                # FilePath1 = str(Path().resolve()) + r'\Excel Files\BSMManagementDailyUpdate.xlsx'
                Sheet2 = 'BSM Page Data'

                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet2, 'Scan Barcode', RowNo)
                RWDE.WriteData(FilePath1, Sheet2, Rowindex1, ColumnNo, BCKId)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1', RowNo)
                RWDE.WriteData(FilePath, Sheet1, Rowindex1, ColumnNo, Tube1)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet2, 'Tube1', RowNo)
                RWDE.WriteData(FilePath1, Sheet2, Rowindex1, ColumnNo, Tube1)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2', RowNo)
                RWDE.WriteData(FilePath, Sheet1, Rowindex1, ColumnNo, Tube2)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet2, 'Tube2', RowNo)
                RWDE.WriteData(FilePath1, Sheet2, Rowindex1, ColumnNo, Tube2)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3', RowNo)
                RWDE.WriteData(FilePath, Sheet1, Rowindex1, ColumnNo, Tube3)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet2, 'Tube3', RowNo)
                RWDE.WriteData(FilePath1, Sheet2, Rowindex1, ColumnNo, Tube3)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4', RowNo)
                RWDE.WriteData(FilePath, Sheet1, Rowindex1, ColumnNo, Tube4)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet2, 'Tube4', RowNo)
                RWDE.WriteData(FilePath1, Sheet2, Rowindex1, ColumnNo, Tube4)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Result', RowNo)
                RWDE.WriteData(FilePath, Sheet1, Rowindex1, ColumnNo, ResultMessage)

                # Go Back to Home Page Button
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[text() = "Go Back to Home Page"]', 60)
                driver.execute_script('arguments[0].click();', Element)

            time.sleep(Seconds)
            if (Rowindex1 == RowCount1):
                # Account Icon
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//p', 60)
                driver.execute_script('arguments[0].click();', Element)

                # Log Out Link
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Log Out"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                time.sleep(3)
                driver.quit()
    else:
        time.sleep(Seconds)
        if (Rowindex1 == RowCount1):
            # Account Icon
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//p', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Log Out Link
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Log Out"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            time.sleep(3)
            driver.quit()
