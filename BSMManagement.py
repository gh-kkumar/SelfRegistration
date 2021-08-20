import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
#import OrderManagement as OM
#import os
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
Url = str(RWDE.ReadData(FilePath, Sheet, 5, ColumnNo))

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-notifications')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver', options=chrome_options)
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\BSMManagement.xlsx'
#FilePath = str(Path().resolve()) + '\Excel Files\BSMManagementDailyUpdate.xlsx'
Sheet = 'BSM Page Data'
RowCount = 2
Seconds = 1

#1. This is for SFDC Login

for RowIndex in range(2, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
    if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "username"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Name')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "pw"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Password')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'RememberMe')
        if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//label[text() = "Remember me"]', 60)
            Element.click()

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Login"]', 60)
        Element.click()

FilePath1 = str(Path().resolve()) + '\Excel Files\PhlebotomistManagement.xlsx'
Sheet1 = 'Phlebotomist page Data'
ColumnNo1 = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'BCK ID', 5)

# BSM Button Div
time.sleep(Seconds)
Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "BSM"]', 60)
driver.execute_script('arguments[0].click();', Element)

RowCount = RWDE.RowCount(FilePath, Sheet)
for Rowindex1 in range(6, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Run', 5)
    if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
        # Scan BarCode TextBox
        time.sleep(2)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "bckbox"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Scan Barcode', 5)
        if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
        else:
            Element.send_keys('')

        # Scan BCK
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[text() = "Scan BCK"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Error Message
        time.sleep(Seconds)
        ErrorExists = WER.check_exists_by_xpath(driver, '/html/body/div[6]/div/div/div')
        # ErrorExists1 = WER.check_exists_by_xpath(driver, '//div[6]/div/div/div/div//span')
        if (ErrorExists == True):
            time.sleep(4)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "bckbox"]', 60)
            Element.clear()
        else:
            # Tube1 TextBox
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb1tube"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Pass Tube1', 5)
            if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1', 5)
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
            else:
                Element.send_keys('')

            ##Tube 1 Volume in ml
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Volume', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb1vol"]',
                                         60)
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

            ## Tube1 Exception
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Exception', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb1excep"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                ddlElement = '(//lightning-base-combobox-item[@data-value = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"])[1]'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, ddlElement, 60)
                driver.execute_script('arguments[0].click();', Element)

            # Tube1 Secondary Check
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Secondary Check', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[@class="slds-checkbox_faux"])[1]', 60)
                Element.click()

            # Tube2 TextBox
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Pass Tube2', 5)
            if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb2tube"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2', 5)
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
            else:
                Element.send_keys('')

            ## Tube 2 Volume in ml
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Volume', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb2vol"]', 60)
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

            ## Tube2 Exception
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Exception', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb2excep"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                ddlElement = '(//lightning-base-combobox-item[@data-value = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"])[2]'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, ddlElement, 60)
                driver.execute_script('arguments[0].click();', Element)

            # Tube2 Secondary Check
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Secondary Check', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[@class="slds-checkbox_faux"])[2]', 60)
                Element.click()

            # Tube3 TextBox
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Pass Tube3', 5)
            if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb3tube"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3', 5)
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
            else:
                Element.send_keys('')

            ## Tube 3 Volume in ml
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3 Volume', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb3vol"]', 60)
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

            ## Tube3 Exception
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3 Exception', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb3excep"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                ddlElement = '(//lightning-base-combobox-item[@data-value = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"])[3]'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, ddlElement, 60)
                driver.execute_script('arguments[0].click();', Element)

            # Tube3 Secondary Check
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3 Secondary Check', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[@class="slds-checkbox_faux"])[3]', 60)
                Element.click()

            # Tube4 TextBox
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Pass Tube4', 5)
            if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb4tube"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4', 5)
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
            else:
                Element.send_keys('')

            ## Tube 4 Volume in ml
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4 Volume', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb4vol"]', 60)
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

            ## Tube 4 Exception
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4 Exception', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "tb4excep"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                ddlElement = '(//lightning-base-combobox-item[@data-value = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"])[4]'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, ddlElement, 60)
                driver.execute_script('arguments[0].click();', Element)

            # Tube4 Secondary Check
            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4 Secondary Check', 5)
            if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[@class="slds-checkbox_faux"])[4]', 60)
                Element.click()

            # Receive Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Receive"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            if (Rowindex1 == RowCount):
                # Account Icon
                time.sleep(3)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//img[@title = "User"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                # Log Out Link
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[text() = "Log Out"]',
                                         60)
                driver.execute_script('arguments[0].click();', Element)

                time.sleep(2)
                driver.quit()
                break
    else:
        if (Rowindex1 == RowCount):
            # Account Icon
            time.sleep(3)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//img[@title = "User"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Log Out Link
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[text() = "Log Out"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            time.sleep(2)
            driver.quit()
            break
    #else:
    #    time.sleep()