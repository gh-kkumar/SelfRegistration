import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import OrderManagement as OM
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
Url = str(RWDE.ReadData(FilePath, Sheet, 3, ColumnNo))


driver = webdriver.Chrome(executable_path = str(Path().resolve()) + r'\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)

#1. This is for HCP/Clinical Staff Login Page

FilePath = str(Path().resolve()) + '\Excel Files\PatientAndOrderManagement.xlsx'
Sheet = 'Patient Information Page Data'

Seconds = 1 #300 / 1000
Seconds1 = 300 / 1000
RowCount = 3
for RowIndex in range(2, RowCount + 1):
    Column = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
    if (RWDE.ReadData(FilePath, Sheet, RowIndex, Column) == 'Y'):
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Name')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Password')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[contains(text(),"Log in")]', 60)
        Element.click()

# New Patient Button
time.sleep(Seconds)
Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "New Patient"]', 60)
driver.execute_script('arguments[0].click();', Element)

RowCount = RWDE.RowCount(FilePath, Sheet)
Phlebotomist_RowNo = 6
for RowIndex1 in range(8, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Run', 6)
    if (RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo) == 'Y'):
        # Patient Information

        # FirstName
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "FirstName"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'First Name', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # LastName
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "LastName"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Last Name', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # Email
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonEmail"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Email', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # Phone
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Phone"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Phone', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'Phone'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # Gender
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "HealthCloudGA__Gender__pc"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Gender Element
        time.sleep(Seconds1)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Gender', 7)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[@title = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # DOB
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonBirthdate"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Birth Date', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # MRN Number
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "MRN_Id__c"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'MRN', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # Address Information
        # Street
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonMailingStreet"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Street', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # City
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonMailingCity"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'City', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # State
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonMailingStateCode"]', 60)
        driver.execute_script('arguments[0].scrollIntoView();', Element)
        Element.click()

        # State Element
        time.sleep(Seconds1)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'State', 7)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[@title = "' + RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo) + '"]', 60)
        driver.execute_script('return arguments[0].scrollIntoView(false);', Element)
        driver.execute_script('arguments[0].style.backgroundColor = "#FAEDEA";', Element)
        Element.click()

        time.sleep(Seconds1)
        # Postal Code
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "PersonMailingPostalCode"]', 60)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Zip Code', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
        else:
            Element.send_keys('')

        # Next Button
        time.sleep(Seconds1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Self Pay
        time.sleep(Seconds)
        ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Self Pay', 7)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) == 'Y'):
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Self Pay"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Next Button
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            OM.Place_Order(driver, FilePath, Sheet, RowIndex1, RowCount, Seconds, Seconds1, Phlebotomist_RowNo, 7, 'New Patient')
        else:
            # Insurance Carrier
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Insurance_Carrier__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Insurance Carrier', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
            else:
                Element.send_keys('')

            # PolicyID
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Policy_Id__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Policy ID', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
            else:
                Element.send_keys('')

            # Insurance Carrier First Name
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Insurance_Carrier_First_Name__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Insurance Carrier First Name', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
            else:
                Element.send_keys('')

            # Insurance Carrier Last Name
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Insurance_Carrier_Last_Name__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Insurance Carrier Last Name', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
            else:
                Element.send_keys('')

            # GroupID
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Group_Id__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Group ID', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))
            else:
                Element.send_keys('')

            # Patient Relation To Insured
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Patient Relation To Insured', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                if(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) == 'Self'):
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Patient_Relation_to_Insured__c"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Patient Relation To Insured Element
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[@title = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"]', 60)
                    driver.execute_script('arguments[0].click();', Element)
                else:
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Patient_Relation_to_Insured__c"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Patient Relation To Insured Element
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[@title = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"])[1]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Gender1
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Gender__c"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Gender1 Element
                    time.sleep(Seconds1)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Gender1', 7)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[@title = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Birth Date1
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Birthdate__c"]', 60)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Birth Date1', 7)
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))

                    # Street1
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Street__c"]', 60)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Street1', 7)
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))

                    # City1
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "City__c"]', 60)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'City1', 7)
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))

                    # State
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "State__c"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # State Element
                    time.sleep(Seconds1)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'State1', 7)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[@title = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Zip Code
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Zipcode__c"]', 60)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Zip Code1', 7)
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))

            # Company Name
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Company_Name__c"]', 60)
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Company Name', 7)
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))

            # Add Insurance
            ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Add Insurance', 7)
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) == 'Y'):
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[text() = "Add Additional Insurance"]', 60)
                driver.execute_script('arguments[0].scrollIntoView(true);', Element)
                Element.click()

                # SecondaryInsuranceCarrier
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Insurance_Carrier__c"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Insurance Carrier', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
                else:
                    Element.send_keys('')

                # SecondaryPolicyID
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Policy_Id__c"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Policy ID', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
                else:
                    Element.send_keys('')

                # SecondaryInsuranceCarrierFirstName
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Insurance_Carrier_First_Name__c"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Insurance Carrier First Name', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
                else:
                    Element.send_keys('')

                # SecondaryInsuranceCarrierLastName
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Insurance_Carrier_Last_Name__c"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Insurance Carrier Last Name', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
                else:
                    Element.send_keys('')

                # SecondaryGroupID
                time.sleep(Seconds1)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Group_Id__c"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Group ID', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)))
                else:
                    Element.send_keys('')

                # SecondaryPatientRelationToInsured
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Patient Relation To Insured', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Patient_Relation_to_Insured__c"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '(//span[text() = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) + '"])[2]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                # Secondary Company Name
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Secondary Company Name', 7)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo)) != 'None'):
                    time.sleep(Seconds1)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@name = "Secondary_Company_Name__c"]', 60)
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo))

            # Next Button
            time.sleep(Seconds1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[@title = "Next"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            time.sleep(4)
            if (WER.check_exists_by_xpath(driver, '/html/body/div[4]/div/div/div') == True):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '/html/body/div[4]/div/div/div', 60)
                ErrorMsg = Element.text.replace('success\n', '').replace('\n', '').replace('error', '').replace('Close', '').replace('success ', '')
                print(ErrorMsg)

                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Result', 6)
                if (RWDE.ReadData(FilePath, Sheet, RowIndex1, ColumnNo) == ErrorMsg and ErrorMsg != 'Record created successfully'):
                    RWDE.WriteData(FilePath, Sheet, RowIndex1, ColumnNo, ErrorMsg)
                    # RWDE.WriteData(FilePath, Sheet, RowIndex1, 30, 'Passed')

                    OM.Skip_From_Flow(driver, RowIndex1, RowCount, Seconds)
                elif (ErrorMsg == 'Record created successfully'):
                    OM.Place_Order(driver, FilePath, Sheet, RowIndex1, RowCount, Seconds, Seconds1, Phlebotomist_RowNo, 7, 'New Patient')
                else:
                    #RWDE.WriteData(FilePath, Sheet, RowIndex1, 29, ErrorMsg)
                    #RWDE.WriteData(FilePath, Sheet, RowIndex1, 30, 'Failed')

                    OM.Skip_From_Flow(driver, RowIndex1, RowCount, Seconds)
            else:
                OM.Place_Order(driver, FilePath, Sheet, RowIndex1, RowCount, Seconds, Seconds1, Phlebotomist_RowNo, 7, 'New Patient')

            Phlebotomist_RowNo += 1