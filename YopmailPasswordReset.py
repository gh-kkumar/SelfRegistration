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

FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 4, 3))

FilePath1 = str(Path().resolve()) + r'\Excel Files\HCPregistration.xlsx'
Sheet1 = 'HCP Registration Page Data'
RowCount = RWDE.RowCount(FilePath1, Sheet1)
for RowIndex in range(2, RowCount + 1):
    if (RWDE.ReadData(FilePath1, Sheet1, RowIndex, RWDE.FindColumnNoByName(FilePath1, Sheet1, 'Result')) == 'Pass'):
        driver = webdriver.Chrome(executable_path=str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver')
        driver.maximize_window()
        driver.get(Url)

        ColumnNo = RWDE.FindColumnNoByName(FilePath1, Sheet1, 'Email')
        Email = RWDE.ReadData(FilePath1, Sheet1, RowIndex, ColumnNo)

        Seconds = 1
        time.sleep(Seconds)
        EmailNameTextBox = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#login')))
        EmailNameTextBox.send_keys(Email)

        time.sleep(Seconds)
        # CheckinboxButton = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//td[3]/input')))
        LoginButton = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//button/i')))
        LoginButton.click()

        driver.switch_to.frame('ifinbox')
        WelcometoLunarHCPSandBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[. = "Sandbox: Welcome to Screening Portal"]')))
        driver.execute_script('arguments[0].click();', WelcometoLunarHCPSandBox)

        driver.switch_to.default_content()
        driver.switch_to.frame('ifmail')
        LunarHcpLink = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//pre')))
        HyperLink = LunarHcpLink.text.replace('\n', '').replace(' ', '')[307: 500]  # .replace('HiJames,WelcometoScreeningPortal!Togetstarted,goto', '').replace('Username:jamesstephen1@yopmail.comThanks,GuardantHealth-ScreeningOrg', '')
        HyperLink = LunarHcpLink.text.replace('\n', '').replace(' ', '').replace(HyperLink, '')
        HyperLink = HyperLink[0: HyperLink.find('https:')]
        HyperLink = LunarHcpLink.text.replace('\n', '').replace(' ', '').replace(HyperLink, '')
        HyperLink = HyperLink.replace(HyperLink[HyperLink.find('Username:'): 500], '')
        # print(HyperLink)

        driver.switch_to.default_content()
        driver.switch_to.window(driver.window_handles[0])
        driver.get(HyperLink)

        ColumnNo = RWDE.FindColumnNoByName(FilePath1, Sheet1, 'Login Password')
        time.sleep(Seconds)
        NewPasswordTextBox = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#newpassword')))
        NewPasswordTextBox.send_keys(RWDE.ReadData(FilePath1, Sheet1, RowIndex, ColumnNo))

        time.sleep(Seconds)
        ConfirmPasswordTextBox = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#confirmpassword')))
        ConfirmPasswordTextBox.send_keys(RWDE.ReadData(FilePath1, Sheet1, RowIndex, ColumnNo))

        time.sleep(Seconds)
        SaveButton = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#password-button')))
        SaveButton.click()

        # Manage Clinic Menu
        Element = '//a[. = "Manage Clinic"]'
        ResultPageMessage = WER.check_exists_by_xpath(driver, Element)
        if (ResultPageMessage == True):
            Text = 'Logged in as Admin,' + driver.title + ' as Page Title ' + 'and Access to "Manage clinic"'
        else:
            Text = 'Logged in as Non Admin,' + driver.title + ' as Page Title ' + 'and No Access to "Manage clinic"'

        ActualResultColumnNo = RWDE.FindColumnNoByName(FilePath1, Sheet1, 'Actual Result After Password Reset')
        RWDE.WriteData(FilePath1, Sheet1, RowIndex, ActualResultColumnNo, Text)

        time.sleep(Seconds)
        AccountImage = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//div/button/div/div[1]/p')))
        AccountImage.click()

        time.sleep(Seconds)
        LogOutLink = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//span[. = "Log Out"]')))
        LogOutLink.click()

        driver.quit()