import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
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


def Place_Order(WebDriver, FileName, Sheet, RowIndex, RowCount, Stime, Stime1, Phlebotomist_RowNo, RequiredRow, Scenario):
    # Order Management Screen
    #RequiredRow = 7
    # Draw Type
    time.sleep(Stime1)
    Element = BEP.WebElement(WebDriver, EC.visibility_of_element_located, By.XPATH, '//input[@name = "Draw_Type__c"]', 60)
    WebDriver.execute_script('arguments[0].click();', Element)

    # Draw Type Element
    ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Draw Type', RequiredRow)
    ElementText = str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo))
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "' + ElementText + '"]', 60)
    WebDriver.execute_script('arguments[0].click();', Element)

    # ICD10 Test CheckBox
    ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'ICD 10', RequiredRow)
    ColumnNo1 = RWDE.FindColumnNoByName1(FileName, Sheet, 'Other', RequiredRow)
    if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) == 'Y'):
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Z12.11 / Z12.12"]', 60)
        if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo1)) == 'Y'):
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Other"]', 60)
            WebDriver.execute_script('arguments[0].click();', Element)

            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Other_ICD_Code__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Other ICD Code', RequiredRow)
            Element.send_keys(str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)))

            # GH Positive Result Disclosure
            time.sleep(Stime1)
            Element = BEP.WebElement(WebDriver, EC.visibility_of_element_located, By.XPATH, '//input[@name = "GH_Positive_Result_Disclosure__c"]', 60)
            WebDriver.execute_script('arguments[0].click();', Element)

            # GHPositiveResult Element
            ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'GH Positive Result Disclosure', RequiredRow)
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "' + str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) + '"]', 60)
            WebDriver.execute_script('arguments[0].click();', Element)
        else:
            # GH Positive Result Disclosure
            time.sleep(Stime1)
            Element = BEP.WebElement(WebDriver, EC.visibility_of_element_located, By.XPATH, '//input[@name = "GH_Positive_Result_Disclosure__c"]', 60)
            WebDriver.execute_script('arguments[0].click();', Element)

            # GHPositiveResult Element
            ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'GH Positive Result Disclosure', RequiredRow)
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "' + str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) + '"]', 60)
            WebDriver.execute_script('arguments[0].click();', Element)

    # Patient Status at Blood Draw
    ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Patient Status at Blood Redraw', RequiredRow)
    if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) != 'None'):
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.visibility_of_element_located, By.XPATH, '//label/span[text() = "' + str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) + '"]', 60)
        WebDriver.execute_script('arguments[0].click();', Element)

    # ShareResult CheckBox
    ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Share Result', RequiredRow)
    if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) == 'Y'):
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Share Result"]', 60)
        Element.click()

        # SecondaryRecipient
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Recipient__c"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Secondary Recipient', RequiredRow)
        if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)))
        else:
            Element.send_keys('')

        # SecondaryRecipientFax
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Recipient_Fax__c"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Secondary Recipient Fax', RequiredRow)
        if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)))
        else:
            Element.send_keys('')

    # Next Button
    time.sleep(Stime1)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
    WebDriver.execute_script('arguments[0].click();', Element)

    ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Place Order', RequiredRow)
    if (str(RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo)) == 'Y'):
        # Submit Button
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Submit"]', 60)
        WebDriver.execute_script('arguments[0].click();', Element)

        # Close Button for QR Code Scan
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Close"]', 60)
        WebDriver.execute_script('arguments[0].click();', Element)

        # Result Message
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '/html/body/div[4]/div/div/div/div/div/span', 60)
        print(Element.text.replace('Order created successfully. Your Order number is ', ''))

        OrderNo = Element.text.replace('Order created successfully. Your Order number is ', '')
        ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Actual Result - OrderID', RequiredRow)
        RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo, OrderNo)

        ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'HCP Approval Required', RequiredRow)
        ColumnNo1 = RWDE.FindColumnNoByName1(FileName, Sheet, 'Order Status', RequiredRow)

        ColumnNo2 = RWDE.FindColumnNoByName(FileName, Sheet, 'Description')
        Description = RWDE.ReadData(FileName, Sheet, RowIndex, ColumnNo2)

        if (Description == 'Clinic Staff'):
            RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo, 'Y')
            RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo1, 'Pending')
        else:
            RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo, 'N')
            RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo1, 'Signed')

        ColumnNo = RWDE.FindColumnNoByName1(FileName, Sheet, 'Update In phlebotomist DataFile', RequiredRow)
        RWDE.WriteData(FileName, Sheet, RowIndex, ColumnNo, 'Y')

        FilePath1 = str(Path().resolve()) + r'\Excel Files\PhlebotomistManagement.xlsx'
        # FilePath1 = str(Path().resolve()) + r'\Excel Files\PhlebotomistManagementDailyUpdate.xlsx'
        Sheet1 = 'Phlebotomist page Data'

        ColumnNo2 = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Run', 5)
        RWDE.WriteData(FilePath1, Sheet1, Phlebotomist_RowNo, ColumnNo2, 'Y')

        ColumnNo2 = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Order Number', 5)
        RWDE.WriteData(FilePath1, Sheet1, Phlebotomist_RowNo, ColumnNo2, OrderNo)

        ColumnNo2 = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Collect Specimen Samples', 5)
        RWDE.WriteData(FilePath1, Sheet1, Phlebotomist_RowNo, ColumnNo2, 'Y')

        if (RowIndex == RowCount):
            time.sleep(Stime)
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//img[@src= "/profilephoto/005/T"]', 60)
            Element.click()

            time.sleep(Stime)
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Log Out"]', 60)
            Element.click()
        else:
            time.sleep(Stime1)
            ElementString = '//button[@title = "'
            if(Scenario == 'Patient Search'):
                ElementString += 'Patient Search"]'
            if(Scenario == 'New Patient'):
                ElementString += 'New Patient"]'

            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, ElementString, 60)
            WebDriver.execute_script('arguments[0].click()', Element)
    else:
        # Popup Cancel Button
        time.sleep(Stime1)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//header/span//*[name()="path"]', 60)
        Element.click()

def Skip_From_Flow(WebDriver, RowIndex, RowCount, Stime):
    time.sleep(Stime)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//header//*[name()="svg"]', 60)
    Element.click()
    if (RowIndex == RowCount):
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//img', 60)
        Element.click()

        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log Out"]', 60)
        Element.click()
    else:
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[. = "New Patient"]', 60)
        WebDriver.execute_script('arguments[0].click()', Element)