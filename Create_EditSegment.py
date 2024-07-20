import fileinput

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

DelayDict ={
    0.5: 10,
    1: 15,
    1.5: 22,
    2: 30,
    2.5: 40,
    3: 48
    }


# //div[@data-test= "name"]//input


def edit_Segment(x12, x1, x3, x4, x5, x6, x7, x9, x10, x11):
    website = "https://studio.iris.microsoft.com/#/segments"

    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(website)
    driver.maximize_window()
    time.sleep(5)

    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(3)

    i = 0
    Counts = []
    for Seg in x1:
        if x12[i] == "Edit":

            driver.get(x9[i] + "/edit")

            SegmentName = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test= "name"]//input')))
            ActionChains(driver).double_click(SegmentName).perform()
            SegmentName.send_keys(x1[i])
            time.sleep(1)

            Description = driver.find_element(by='xpath', value='//div[@data-test= "description"]//textarea')

            ActionChains(driver).double_click(Description).perform()
            Description.send_keys(Keys.CONTROL+"A")
            Description.send_keys(x3[i])
            time.sleep(1)

            DelayInPublish = driver.find_element(by='xpath', value='//div[@data-test= "delayInHours"]//span[contains(@class,"Slider-thumb")]')
            ActionChains(driver).drag_and_drop_by_offset(DelayInPublish, -400, 0).perform()
            time.sleep(1)
            ActionChains(driver).drag_and_drop_by_offset(DelayInPublish, 10, 0).perform()

            try:
                TagsList = driver.find_element(by='xpath', value='//div[@data-test="tags"]/div/div/div/div[2]/div/div').get_attribute("data-selection-index")
                TagsList = int(TagsList)
            except NoSuchElementException:
                TagsList = 1

            while TagsList == 0:
                Tag = driver.find_element(by='xpath', value='//div[@data-test="tags"]/div/div/div/div[2]/div/div/span[2]')
                Tag.click()
                time.sleep(1)
                try:
                    TagsList = driver.find_element(by='xpath', value='//div[@data-test="tags"]/div/div/div/div[2]/div/div').get_attribute("data-selection-index")
                    TagsList = int(TagsList)
                except NoSuchElementException:
                    break
            time.sleep(1)

            CampTag = driver.find_element(by='xpath', value='//div[@data-test= "tags"]//input')
            CampTag.send_keys(x5[i])
            Group = driver.find_element(by='xpath', value='//div[@data-test= "groups"]//input')
            Group.click()
            CampTag.send_keys(x6[i])
            Group.click()
            CampTag.send_keys(x7[i])
            Group.click()
            NextButton = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[2]')
            NextButton.click()
            time.sleep(1)

            TextArea = driver.find_element(by='xpath', value='//div[@id="customQuery"]/textarea')
            ActionChains(driver).double_click(TextArea).perform()
            TextArea.send_keys(Keys.CONTROL+"A")
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

            NextButton2 = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[3]')
            NextButton2.click()
            time.sleep(1)

            NextButton3 = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            NextButton3.click()
            time.sleep(1)

            SubmitButton = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            SubmitButton.click()
            time.sleep(1)

            SegmentLink = WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.XPATH, '//div[(@data-automationid="DetailsRowFields")]/div[(@aria-colindex="1")]/a')))

            Counts.append(i)
            i = i+1
        else:
            i = i+1
            continue
    driver.quit()
    return Counts