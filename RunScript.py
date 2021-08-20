import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import SelfRegistration as SR
import Login as Log
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

FilePath = str(Path().resolve()) + r'\Excel Files\RunScript.xlsx'
Sheet = 'Run Scenario'
RowCount = RWDE.RowCount(FilePath, Sheet)
ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
ColumnNo1 = RWDE.FindColumnNoByName(FilePath, Sheet, 'Scenario')
for RowIndex in range(2, RowCount + 1):
   if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
        if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo1) == 'SelfRegistration'):
            SR.HCPOrClinicStaff_SelfRegistration()
        if (RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo1) == 'Login'):
            Log.HCPOrClinicalStaff_Login()
