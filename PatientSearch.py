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

FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 3, 3))

driver = webdriver.Chrome(executable_path = str(Path().resolve()) + r'\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)

#1. This is for HCP Login Page

FilePath = str(Path().resolve()) + r'\Excel Files\PatientSearch.xlsx'
Sheet = 'Login Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)
Seconds = 300 / 1000

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log in"]', 60)
    Element.click()

    time.sleep(7)
    print(driver.title)
    if (driver.title == 'Login'):
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div/div/span/div/div', 60)
        if(RWDE.ReadData(FilePath, Sheet, RowIndex, 5) == Element.text):
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, Element.text)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Passed')
        else:
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, Element.text)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Failed')

        driver.execute_script('arguments[0].innerHTML = ""', Element)

        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
        Element.clear()

        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
        Element.clear()
    elif (driver.title == 'Guardant Health'):
        if (RWDE.ReadData(FilePath, Sheet, RowIndex, 5) == driver.title):
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, driver.title)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Passed')

        # 2. This is for Patient Search Flow
        Sheet = 'Patient Search Page Data'
        RowCount = RWDE.RowCount(FilePath, Sheet)

        # Patient Search
        time.sleep(1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Patient Search"]', 60)
        Element.click()

        for RowIndex1 in range(3, RowCount + 1):
            # First Name
            time.sleep(1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div/lightning-input//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 2)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 2)))
            else:
                Element.clear()
                Element.click()

            # Last Name
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div/lightning-input//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 3)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 3)))
            else:
                Element.clear()
                Element.click()

            # MRN
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/lightning-input//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 4)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 4)))
            else:
                Element.clear()
                Element.click()

            # DOB
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div/div/input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 5)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 5)))
            else:
                Element.clear()
                Element.click()

            # Phone
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[5]//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 6)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 6)))
            else:
                Element.clear()
                Element.click()

            # Search Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Search"]', 60)
            Element.click()
            #print(RowIndex1)
            #if (WER.check_exists_by_xpath(driver, '/html/body/div[4]/div/div/div')):
            #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '/html/body/div[4]/div/div/div', 60)
            #    Text = driver.execute_script('return arguments[0].innerText', Element).replace('error\n', 'error ').replace('\nClose', '')
            #    if(RowIndex1 != RowCount):
            #        RWDE.WriteData(FilePath, Sheet, RowIndex1, 14, Text)
            #    else:
            #        RWDE.WriteData(FilePath, Sheet, RowIndex1 + 1, 14, Text)

            #    print(Text)

            time.sleep(2)
            # Table
            Element = '//div[2]/div/div/table'
            TableRows = Element + '/tbody/tr'
            row_count = len(driver.find_elements_by_xpath(TableRows))
            CellVal = ['', '', '', '']
            if(row_count > 0) :
                for RIndex in range(1, row_count + 1):
                    TableColumns = driver.find_elements_by_xpath(Element + '/tbody/tr[' + str(RIndex) + ']/td')
                    TableCellValue1 = driver.find_element(By.XPATH, '//tbody/tr[' + str(RIndex) + ']/th')
                    CellVal[0] = TableCellValue1.text
                    for CIndex in range(1, len(TableColumns) + 1):
                        TableCellValue = driver.find_element_by_xpath(Element + '/tbody/tr[' + str(RIndex) + ']/td[' + str(CIndex) + ']')
                        CellVal[CIndex] = TableCellValue.text
                        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 9)) == 'None'):
                            CellValue = ''
                        else:
                            CellValue = RWDE.ReadData(FilePath, Sheet, RowIndex1, 9)

                    if (CellVal[0] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 7) and
                            CellVal[1] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 8) and
                            str(CellVal[2]) == CellValue):
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 10, CellVal[0])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 11, CellVal[1])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 12, CellVal[2])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 13, 'Passed')
                    else:
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 13, 'Failed')
                    break
            else:
                RWDE.WriteData(FilePath, Sheet, RowIndex1, 13, 'Passed')

















