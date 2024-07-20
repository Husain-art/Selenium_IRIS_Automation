from selenium import webdriver
from selenium.webdriver import Keys
from selenium .webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.edge.service import Service

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import NoSuchElementException

from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from selenium.webdriver.remote import webelement
from selenium.webdriver.support.wait import WebDriverWait

DelayDict = {0.5: 10, 1: 15, 1.5: 22, 2: 30, 2.5: 40, 3: 48}


# //div[@data-test= "name"]//input


def create_Segment(x12, x1, x2, x3, x4, x5, x6, x7, x8, x10, x11):
    website = 'https://studio.iris.microsoft.com/#/segments/create'

    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(website)
    driver.maximize_window()

    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(3)
    i = 0
    x9 = []
    Count = []
    for Seg in x1:
        if x12[i] == "Create":
            driver.get('https://studio.iris.microsoft.com/#/segments/create')

            SegmentName = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test= "name"]//input')))
            SegmentName.send_keys(x1[i])

            Product = driver.find_element(by='xpath', value='//div[@data-test= "product"]/div/div')
            Product.click()

            ProductValue_XpathText = '//button[@title=\"'+x2[i]+'\"]'
            # print(ProductValue_XpathText)
            ProductButton = driver.find_element(by='xpath', value=ProductValue_XpathText)
            ProductButton.click()

            Description = driver.find_element(by='xpath', value='//div[@data-test= "description"]//textarea')
            Description.send_keys(x3[i])

            # DelayDict[x4[i]]
            time.sleep(2)
            DelayInPublish = driver.find_element(by='xpath', value='//div[@data-test= "delayInHours"]//span[contains(@class,"Slider-thumb")]')
            ActionChains(driver).drag_and_drop_by_offset(DelayInPublish, 10, 0).perform()
            time.sleep(2)

            CampTag = driver.find_element(by='xpath', value='//div[@data-test= "tags"]//input')
            CampTag.send_keys(x5[i])
            Group = driver.find_element(by='xpath', value='//div[@data-test= "groups"]//input')
            Group.click()
            CampTag.send_keys(x6[i])
            Group.click()
            CampTag.send_keys(x7[i])
            Group.click()
            # //main/div/div/div[2]/div/div/form/div[2]/button[2]
            NextButton = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[2]')
            NextButton.click()

            # name="Custom Scope Query"
            CustomQueryButton = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//button[@name= "Custom Scope Query"]')))
            CustomQueryButton.click()

            # //button[@data-test="PositiveDialogButton"]
            PositiveResp = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//button[@data-test="PositiveDialogButton"]')))
            PositiveResp.click()

            # //div[@data-test="userIdType"]/div/div
            UserTypeButton = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="userIdType"]/div/div')))
            UserTypeButton.click()

            # //button[@aria-label="PUIDINT"]
            UserType_XpathText = '//button[@aria-label=\"'+x8[i]+'\"]'
            UserType = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, UserType_XpathText)))
            UserType.click()

            # //div[@id="customQuery"]/textarea
            TextArea = driver.find_element(by='xpath', value='//div[@id="customQuery"]/textarea')
            ActionChains(driver).double_click(TextArea).perform()
            TextArea.send_keys(Keys.CONTROL + "A")
            # infile = open(r"C:\Users\v-hcyclewala\OneDrive - Microsoft\V Desktop\PythonHandson\DummyScript.txt", 'r')
            # script = infile.read()
            Filepath = x11
            fileName = str(x10[i])
            Inputpath = Filepath + fileName + ".txt"
            with open(Inputpath) as f:
                lines = f.readlines()

            TotalLines = len(lines)
            for line in lines:
                TextArea.send_keys(line)
                TextArea.send_keys(Keys.HOME)
            f.close()

            # //main/div/div/div[2]/div/div/form/div[2]/button[3]
            NextButton2 = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[3]')
            NextButton2.click()
            time.sleep(2)

            NextButton3 = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            NextButton3.click()
            time.sleep(2)

            SubmitButton = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            SubmitButton.click()
            time.sleep(5)

            #//div[(@data-automationid="DetailsRowFields")]/div[(@aria-colindex="1")]/a
            SegmentLink = WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.XPATH, '//div[(@data-automationid="DetailsRowFields")]/div[(@aria-colindex="1")]/a')))
            x9.append(SegmentLink.get_attribute('href'))
            Count.append(i)
            i = i+1
        else:
            i = i+1
            continue
        time.sleep(3)
    driver.quit()
    return x9, Count

    # driver.quit()
    # //span[@aria-label = "ID Field"]
    # time.sleep(8)