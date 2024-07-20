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

def create_Interaction(x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23):
    x24 = []
    counts = []
    website = x22[0]
    print(x22)

    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(website)
    driver.maximize_window()

    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(3)
    i = 0
    for CrtInt in x22:
        if x23[i] == "Create":
            AddInteractionLink = CrtInt + "/interactions/add"
            driver.get(AddInteractionLink)

            InteractionName = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="interactions"]//div[@data-test= "name"]//input')))
            InteractionName.send_keys(x6[i])

            # //div[@data-test="interactions"]//div[@data-test= "description"]//textarea
            InteractionDescription = driver.find_element(by='xpath', value='//div[@data-test="interactions"]//div[@data-test= "description"]//textarea')
            InteractionDescription.send_keys(x7[i])

            # # SelectOucome start
            # //h1[text() = 'Select outcomes']
            try:
                WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "//h1[text()='Select outcomes']")))
                action = ActionChains(driver)
                # NoOutcome = driver.find_element(by='xpath',
                #                                        value="//span[text()='No Outcome']//parent::div//parent::div//parent::div")

                Outcome = WebDriverWait(driver, 120).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@data-test="outcomes"]//div[@data-list-index="19"]')))
                action.scroll_to_element(Outcome).perform()

                FocusedProduct = driver.find_element(by='xpath', value='//div[@data-test="productList"]')
                SelectedItemsResult = driver.find_element(by='xpath', value='//div[@data-test="outcomes"][2]')

                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()
                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()

                Outcome = WebDriverWait(driver, 120).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@data-test="outcomes"]//div[@data-list-index="39"]')))
                action.scroll_to_element(Outcome).perform()
                Outcome.click()
                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()
                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()

                Outcome = WebDriverWait(driver, 120).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@data-test="outcomes"]//div[@data-list-index="59"]')))
                action.scroll_to_element(Outcome).perform()
                Outcome.click()

                # action.scroll_to_element(FocusedProduct).perform()
                # action.scroll_to_element(SelectedItemsResult).perform()
                # action.scroll_to_element(FocusedProduct).perform()
                # action.scroll_to_element(SelectedItemsResult).perform()
                #
                # Outcome = WebDriverWait(driver, 120).until(EC.presence_of_element_located(
                #     (By.XPATH, '//div[@data-test="outcomes"]//div[@data-list-index="69"]')))
                # action.scroll_to_element(Outcome).perform()
                # Outcome.click()

                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()
                action.scroll_to_element(FocusedProduct).perform()
                action.scroll_to_element(SelectedItemsResult).perform()

                Outcome = WebDriverWait(driver, 120).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@data-test="outcomes"]//div[@data-list-index="54"]')))
                action.scroll_to_element(Outcome).perform()
                Outcome.click()
                # Select Outcome End
                time.sleep(1)
            except NoSuchElementException:
                print("Select Outcome Section is not present")
            # #Select Outcome End
            time.sleep(1)

            NextButton = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[2]')
            NextButton.click()

            UserType = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="identityType"]/div/div')))
            UserType.click()

            UserTypeValue_XpathText = '//button[@title=\"' + x8[i] + '\"]'
            # print(ProductValue_XpathText)
            UserTypeButton = driver.find_element(by='xpath', value=UserTypeValue_XpathText)
            UserTypeButton.click()

            SegmentnameButton = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-item-key="name"]')))
            SegmentnameButton.click()

            SegmentNameInput = driver.find_element(by='xpath', value='//input[@placeholder="Filter by Segment Name"]')
            SegmentNameInput.send_keys(x9[i])

            Filter = driver.find_element(by='xpath', value='//div/div/div/div/div/div/div/ul/li[7]//button[2]')
            Filter.click()

            IncludeButton = driver.find_element(by='xpath', value='//div[@data-automation-key="include"]//button')
            IncludeButton.click()
            time.sleep(1)

            NextButton2 = driver.find_element(by='xpath', value='//main/div/div/div[2]/div/div/form/div[2]/button[3]')
            NextButton2.click()

            Control = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="controlGroupSize"]//input')))
            Control.send_keys(x10[i])

            TreatmentName = driver.find_element(by='xpath', value='//div[@data-automation-key="name"]//input')
            TreatmentName.send_keys(Keys.CONTROL + "A")
            time.sleep(1)
            TreatmentName.send_keys(x11[i])

            time.sleep(1)

            NextButton2.click()

            ActionButton = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//button[@data-test="AddActionIcon"]')))
            ActionButton.click()

            Action_Xpath = '//button[@name=\"' + x12[i] + '\"]'
            # print(ProductValue_XpathText)
            ActionTake = driver.find_element(by='xpath', value=Action_Xpath)
            ActionTake.click()
            # //div[@data-test="platformType"]//div[@tabindex="0"]
            PlatformType = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="platformType"]')))
            PlatformType.click()

            # //button[@title="SFMC Data Extension"]
            # //button[contains(@title, "SFMC")]
            # //button[contains(@title, "Adobe")]
            time.sleep(.5)
            PlatformTypeValue_Xpath = '//button[contains(@title,\"' + str(x13[i]) + '\")]'
            # print(ProductValue_XpathText)
            PlatformTypeValue = driver.find_element(by='xpath', value=PlatformTypeValue_Xpath)
            PlatformTypeValue.click()

            # //button[@name="Email"]
            # //div[@data-test="platformType"]//span
            # //button[@title="SFMC Data Extension"]
            # //button[@title="Iris SMTP"]
            # title="Azure Email Orchestrator"
            # title="Adobe Campaign"

            InstanceName = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="instanceName"]//span')))
            InstanceName.click()

            InstanceNameValue_Xpath = '//button[@title=\"' + x14[i] + '\"]'
            # print(ProductValue_XpathText)
            InstanceNameValue = driver.find_element(by='xpath', value=InstanceNameValue_Xpath)
            InstanceNameValue.click()

            # data-test="instanceName"
            # //button[@title="SFMC_DEVICES_STUDIOS_XBOX_US_PROD"]

            TopicId = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="topicId"]//span')))
            TopicId.click()
            # Xbox Promotional Communications
            TopicIdValue_Xpath = '//button[@title=\"' + x15[i] + '\"]'
            # print(ProductValue_XpathText)
            TopicIdValue = driver.find_element(by='xpath', value=TopicIdValue_Xpath)
            TopicIdValue.click()

            Interval = driver.find_element(by='xpath', value='//div[@data-test="interval"]//input')
            Interval.send_keys(Keys.CONTROL + "A")
            time.sleep(.5)
            Interval.send_keys(x16[i])
            # //div[@data-test="interval"]//input

            NextButton2.click()

            StartDate = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="startDate"]//input[@placeholder="Choose date"]')))
            StartDate.click()
            time.sleep(.5)
            year = x17[i].strftime("%Y")
            # print(year)
            IRISYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[contains(@class, "ms-DatePicker-header header")]/div').text
            # print("Before Increament", IRISYear)

            while IRISYear != year:
                ClickNextYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-yearComponents")]//button[contains(@aria-label, "Go to next year")]')
                ClickNextYear.click()
                time.sleep(.5)
                IRISYear = driver.find_element(by='xpath', value='//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[contains(@class, "ms-DatePicker-header header")]/div').text
                # print(IRISYear)
            # print("AfterIncreament", IRISYear)

            month = x17[i].month
            import calendar
            MonthName = calendar.month_name[month]
            # print("MonthName", MonthName)

            Month_Xpath = '//div[contains(@class, "ms-DatePicker-monthPicker monthPicker")]//div[3]//button[contains(@aria-label, \"' + MonthName + '\")]'
            # print(ProductValue_XpathText)
            ClickMonth = driver.find_element(by='xpath', value=Month_Xpath)
            ClickMonth.click()
            time.sleep(.5)

            Day = x17[i].day
            # print("Day", Day)

            Month_Day = MonthName + " " + str(Day)
            # print(Month_Day)

            # //td/div[contains(@aria-label, "December 15")]
            Day_Xpath = '//td/div[contains(@aria-label, \"' + Month_Day + '\")]'
            # print(Day_Xpath)
            ClickDay = driver.find_element(by='xpath', value=Day_Xpath)
            ClickDay.click()
            time.sleep(.5)

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
            time.sleep(10)

            InteractionLink = WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.XPATH, '//div[@data-automation-key="name"]/a')))
            x24.append(InteractionLink.get_attribute('href'))
            counts.append(i)
            print(x6[i], "Created")
            i = i+1
        else:
            i = i + 1
            continue
            # title="1"
        time.sleep(30)
    driver.quit()
    return x24, counts
