# Iris Studio (microsoft.com)
# //a[@data-test="addInteraction"]

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
from openpyxl import load_workbook
from selenium.webdriver.remote import webelement
from selenium.webdriver.support.wait import WebDriverWait


# //div[@data-test= "name"]//input
IRISError_Xpath = '//div[@data-test="Error report"]'

def Reschedule_Interaction(InteractionLinks, x18, x19, x20, x21, Task):
    website = InteractionLinks[0]
    print(website)

    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(website)
    driver.maximize_window()

    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(3)
    i = 0
    for RSchdInt in InteractionLinks:
        if Task[i] == "Reschedule":
            RescheduleInteractionLink = RSchdInt + "/edit"
            driver.get(RescheduleInteractionLink)

            time.sleep(10)
            NextButton = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[2]')))
            NextButton.click()

            time.sleep(10)
            NextButton2 = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            NextButton2.click()

            time.sleep(2)
            NextButton2 = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            NextButton2.click()

            time.sleep(2)
            NextButton2 = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/div/form/div[2]/button[3]')))
            NextButton2.click()

            DateStr = str(x18[i])
            Dt = DateStr
            from datetime import datetime

            # d = datetime.strptime(DateStr, "%H:%M:%S")
            # Dt = d.strftime("%I:%M %p")

            Hours = Dt[0:2]
            # print(Hours)
            Min = Dt[3:5]
            # print(Min)
            AMPM = Dt[6:8]
            # print(AMPM)

            StartTime = driver.find_element(by='xpath', value='//div[@data-test="time"]//div[@data-test="startDate"]//input')
            StartTime.click()
            StartTime.send_keys(Hours)
            StartTime.send_keys(Min)
            StartTime.send_keys(AMPM)

            ScheduleType = driver.find_element(by='xpath', value='//div[@data-test="scheduleType"]//span')
            ScheduleType.click()

            scheduleTypeValue_Xpath = '//button[@title=\"' + x19[i] + '\"]'
            # print(ProductValue_XpathText)
            scheduleTypeValue = driver.find_element(by='xpath', value=scheduleTypeValue_Xpath)
            scheduleTypeValue.click()

            # /////////////Setting End Date///////////////////////
            EndDate = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="endDate"]//input[@placeholder="Choose date"]')))
            EndDate.click()
            time.sleep(0.5)
            year = x20[i].strftime("%Y")
            # print(year)
            IRISYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[contains(@class, "ms-DatePicker-header header")]/div').text
            # print("Before Increament", IRISYear)

            while IRISYear != year:
                ClickNextYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-yearComponents")]//button[contains(@aria-label, "Go to next year")]')
                ClickNextYear.click()
                time.sleep(.5)
                IRISYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[contains(@class, "ms-DatePicker-header header")]/div').text
            # print("AfterIncreament", IRISYear)

            month = x20[i].month
            import calendar
            MonthName = calendar.month_name[month]
            # print("MonthName", MonthName)

            Month_Xpath = '//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[3]//button[contains(@aria-label, \"' + MonthName + '\")]'
            # print(ProductValue_XpathText)
            ClickMonth = driver.find_element(by='xpath', value=Month_Xpath)
            ClickMonth.click()
            time.sleep(.5)

            Day = x20[i].day
            # print("Day", Day)

            Month_Day = MonthName + " " + str(Day)
            # print(Month_Day)

            # //td/div[contains(@aria-label, "December 15")]
            Day_Xpath = '//td/div[contains(@aria-label, \"' + Month_Day + '\")]'
            # print(Day_Xpath)
            ClickDay = driver.find_element(by='xpath', value=Day_Xpath)
            ClickDay.click()
            time.sleep(.5)

            DayOfMonth = driver.find_element(by='xpath', value='//div[@data-test="dayOfTheMonth"]')
            DayOfMonth.click()

            DayValue_Xpath = '//button[@title=\"' + str(x21[i]) + '\"]'
            # print(DayValue_Xpath)
            DayValue = driver.find_element(by='xpath', value=DayValue_Xpath)
            DayValue.click()
            time.sleep(.5)

            NextButton2.click()
            time.sleep(10)
            NextButton2.click()
            AddInteractionLink = WebDriverWait(driver, 360).until(EC.presence_of_element_located(
                (By.XPATH, '//a[@data-test="addInteraction"]')))

            print(InteractionLinks[i], "Rescheduled")
            i = i+1
        else:
            i = i + 1
            continue
            # title="1"
        time.sleep(30)
    driver.quit()
