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


def HCPOrClinicStaff_SelfRegistration():
    # 1. Registering New HCP in LunarHCPPortalV1

    FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
    Sheet = 'Portal Urls'
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Url')
    Url = str(RWDE.ReadData(FilePath, Sheet, 3, ColumnNo))

    driver = webdriver.Chrome(executable_path=str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver')
    driver.maximize_window()
    driver.get(Url)

    FilePath = str(Path().resolve()) + '\Excel Files\HCPregistration.xlsx'

    Seconds = 1
    Sheet = 'HCP Registration Page Data'
    RowCount = RWDE.RowCount(FilePath, Sheet)
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Actual Result  After Self Registration')
    ColumnNo1 = RWDE.FindColumnNoByName(FilePath, Sheet, 'Result')

    # time.sleep(Seconds)
    # LoginButton = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log in"]', Seconds)
    # LoginButton.click()

    for RowIndex in range(2, RowCount + 1):
        ColumnNo2 = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
        if (RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo2) == 'Y'):
            time.sleep(Seconds)
            Notamemberlink = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Not a member?"]', 60)
            Notamemberlink.click()
            # This is for filling data in the fields
            time.sleep(Seconds)
            FirstNameTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                              '//div[2]/lightning-input//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'First Name')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                FirstNameTextBox.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)))
            else:
                FirstNameTextBox.click()

            time.sleep(Seconds)
            LastNameTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                             '//div[3]/lightning-input//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Last Name')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                LastNameTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                LastNameTextBox.click()

            time.sleep(Seconds)
            EmailTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[9]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Email')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                EmailTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                EmailTextBox.click()

            time.sleep(Seconds)
            StreetTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                           '//div[15]/lightning-input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Street')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                StreetTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                StreetTextBox.click()

            time.sleep(Seconds)
            UserTypeDropDownList = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[22]//select',
                                                  60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Type')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                SelectUserTypeDropDownList = Select(UserTypeDropDownList)
                UserTypeDropDownListOptions = SelectUserTypeDropDownList.options
                SelectUserTypeDropDownList.select_by_visible_text(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                UserTypeDropDownList.click()

            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Site Admin')
            if RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y':
                SiteAdminCheckBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                   '//div[23]/label/span/span[1]', 60)
                SiteAdminCheckBox.click()

            time.sleep(Seconds)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Type')
            ColumnNo1 = RWDE.FindColumnNoByName(FilePath, Sheet, 'NPI Number')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'Provider'):
                NPITextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[29]//input', 60)
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo1)) != 'None'):
                    NPITextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo1))
                else:
                    NPITextBox.click()

            time.sleep(Seconds)
            ClinicTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[5]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Clinic Name')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                ClinicTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                ClinicTextBox.click()

            time.sleep(Seconds)
            PhoneTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[11]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Phone Number')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                PhoneTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                PhoneTextBox.click()

            time.sleep(Seconds)
            CityTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[17]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'City')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                CityTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                CityTextBox.click()

            time.sleep(Seconds)
            StateDropDownList = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[18]//select',
                                               60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'State')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                SelectStateDropDownList = Select(StateDropDownList)
                StateDropDownListOptions = SelectStateDropDownList.options
                SelectStateDropDownList.select_by_visible_text(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                StateDropDownList.click()

            time.sleep(Seconds)
            ZipCodeTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[19]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Zip Code')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                ZipCodeTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                ZipCodeTextBox.click()

            time.sleep(Seconds)
            FaxTextBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[25]//input', 60)
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Fax')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) != 'None'):
                FaxTextBox.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))
            else:
                FaxTextBox.click()

            # time.sleep(Seconds)
            # if RWDE.ReadData(FilePath, Sheet, RowIndex, 15) == 'S':
            #    PhlebotomistCheckBox = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[31]/label/span/span[1]', 60)
            #    PhlebotomistCheckBox.click()

            time.sleep(Seconds)
            SubmitButton = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Submit"]',
                                          60)
            SubmitButton.click()

            # This is for catching validation errors
            ErrorMsg = ''
            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'First Name')
            ColumnNo3 = RWDE.FindColumnNoByName(FilePath, Sheet, 'S.No')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                FirstNameErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                   '//div[. = "First name is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + FirstNameErrorMsg.text)
                ErrorMsg = FirstNameErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Last Name')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                LastNameErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                  '//div[. = "Last name is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + LastNameErrorMsg.text)
                ErrorMsg += ' ' + LastNameErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Email')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                EmailErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                               '//div[. = "Email is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + EmailErrorMsg.text)
                ErrorMsg += ' ' + EmailErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Street')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                StreetErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                '//div[. = "Street is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + StreetErrorMsg.text)
                ErrorMsg += ' ' + StreetErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'NPI Number')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                NPIErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                             '//div[. = "NPI is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + NPIErrorMsg.text)
                ErrorMsg += ' ' + NPIErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Clinic Name')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                ClinicErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                '//div[. = "Hospital/Clinic is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + ClinicErrorMsg.text)
                ErrorMsg += ' ' + ClinicErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Phone Number')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                PhoneErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                               '//div[. = "Phone is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + PhoneErrorMsg.text)
                ErrorMsg += ' ' + PhoneErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'City')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                CityErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                              '//div[. = "City is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + CityErrorMsg.text)
                ErrorMsg += ' ' + CityErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'State')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                StateErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                               '//div[. = "State is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + StateErrorMsg.text)
                ErrorMsg += ' ' + StateErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Zip Code')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                ZipCodeErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                                 '//div[. = "Zip Code is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + ZipCodeErrorMsg.text)
                ErrorMsg += ' ' + ZipCodeErrorMsg.text

            ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Fax')
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo)) == 'None'):
                FaxErrorMsg = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH,
                                             '//div[. = "Fax is mandatory!"]', 60)
                print(str(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo3)) + ' : ' + FaxErrorMsg.text)
                ErrorMsg += ' ' + FaxErrorMsg.text

            time.sleep(8)
            Element = '//div[3]/div/div/div[3]/div'
            ResultPageMessage = WER.check_exists_by_xpath(driver, Element)
            if (ResultPageMessage == True):
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, Element, 60)
                ActualResultColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet,
                                                               'Actual Result  After Self Registration')
                ResultColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Result')
                ExpectedResultColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet,
                                                                 'Expected Result  After Self Registration')
                if (RWDE.ReadData(FilePath, Sheet, RowIndex, ExpectedResultColumnNo) == Element.text):
                    if (ErrorMsg != ''):
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ActualResultColumnNo, ErrorMsg + ' ' + Element.text)
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ResultColumnNo, 'Fail')
                    else:
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ActualResultColumnNo, Element.text)
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ResultColumnNo, 'Pass')
                else:
                    if (ErrorMsg != ''):
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ActualResultColumnNo, ErrorMsg + ' ' + Element.text)
                    else:
                        RWDE.WriteData(FilePath, Sheet, RowIndex, ActualResultColumnNo, Element.text)
                    RWDE.WriteData(FilePath, Sheet, RowIndex, ResultColumnNo, 'Fail')

            LoginButton = BEP.WebElement(driver, EC.visibility_of_element_located, By.XPATH, '//button[. = "Log in"]',
                                         60)
            LoginButton.click()

    time.sleep(3)
    driver.quit()