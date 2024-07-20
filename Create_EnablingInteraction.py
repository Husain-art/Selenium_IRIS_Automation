from selenium import webdriver
from selenium.webdriver import Keys
from selenium .webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.edge.service import Service

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

import time
import pandas as pd
from selenium.webdriver.remote import webelement
from selenium.webdriver.support.wait import WebDriverWait
# pasting file from IRIS Segment
def EnableInteraction(x6, CampaignLink, Task):
    website = CampaignLink
    counts = []
    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(website)
    driver.maximize_window()

    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(3)
    i = 0
    action = ActionChains(driver)
    for Interaction in x6:
        ApprovalStatexpath = '//div[@data-test=\"' +Interaction+'\"]//div[@data-test="approvalState"]/span'
        ApprovalState = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ApprovalStatexpath))).text
        InteractionStateXpath = '//div[@data-test=\"' + Interaction +'\"]//div[@data-test="interactionState"]/span'
        InteractionState = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, InteractionStateXpath))).text


        ActionButtonXpath = '//div[@data-test=\"' +Interaction+'\"]//button'
        print(Interaction)
        if(Task[i] == "Enable" and ApprovalState=="None" and (InteractionState=="Draft" or InteractionState=="ReDraft")):
            ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
            time.sleep(3)
            action.scroll_to_element(ActionButton).perform()
            time.sleep(3)
            ActionButton.click()
            time.sleep(3)
            print("ActionButton.click()")

            RequestApproval = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Request Approval"]')))
            RequestApproval.click()
            time.sleep(2)
            print('RequestApproval.click()')

            YesButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//body/div/div/div/div/div[2]/div/div[2]/div[2]//button')))
            YesButton.click()
            print('YesButton.click()')
            time.sleep(8)

            ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
            time.sleep(3)
            action.scroll_to_element(ActionButton).perform()
            time.sleep(3)
            ActionButton.click()
            print("ActionButton.click()")
            time.sleep(2)

            Approve = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Approve"]')))
            Approve.click()
            print('Approve.click()')
            time.sleep(4)

            attempt = 0
            while attempt < 5:
                try:
                    ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
                    time.sleep(3)
                    action.scroll_to_element(ActionButton).perform()
                    time.sleep(3)
                    ActionButton.click()
                    break
                except StaleElementReferenceException:
                    driver.refresh()
                attempt = attempt + 1
            print("ActionButton.click()")
            time.sleep(2)

            EnableButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Enable"]')))
            EnableButton.click()
            print('EnableButton.click()')
            time.sleep(2)

            YesButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//body/div/div/div/div/div[2]/div/div[2]/div[2]//button')))
            YesButton.click()
            print('YesButton.click()')
            time.sleep(8)

            counts.append(i)
            print(x6[i], "Enabled")
            i = i + 1

        elif(Task[i]  == "Enable" and ApprovalState=="Requested" and (InteractionState=="Draft"  or InteractionState=="ReDraft")):
            ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
            time.sleep(3)
            action.scroll_to_element(ActionButton).perform()
            time.sleep(3)
            ActionButton.click()
            print("ActionButton.click()")
            time.sleep(2)

            Approve = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Approve"]')))
            Approve.click()
            print('Approve.click()')
            time.sleep(4)


            attempt = 0
            while attempt < 5:
                try:
                    ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
                    time.sleep(3)
                    action.scroll_to_element(ActionButton).perform()
                    time.sleep(3)
                    ActionButton.click()
                    break
                except StaleElementReferenceException:
                    driver.refresh()
                attempt = attempt + 1
            print("ActionButton.click()")
            time.sleep(2)

            EnableButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Enable"]')))
            EnableButton.click()
            print('EnableButton.click()')
            time.sleep(2)

            YesButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//body/div/div/div/div/div[2]/div/div[2]/div[2]//button')))
            YesButton.click()
            print('YesButton.click()')
            time.sleep(8)

            counts.append(i)
            print(x6[i], "Enabled")
            i = i + 1

        elif (Task[i]  == "Enable" and ApprovalState == "Approved" and (InteractionState == "Draft" or InteractionState == "Paused" or InteractionState == "ReDraft")):
            ActionButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, ActionButtonXpath)))
            time.sleep(3)
            action.scroll_to_element(ActionButton).perform()
            time.sleep(3)
            ActionButton.click()
            print("ActionButton.click()")
            time.sleep(2)

            EnableButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//button[@name="Enable"]')))
            EnableButton.click()
            print('EnableButton.click()')
            time.sleep(2)

            YesButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//body/div/div/div/div/div[2]/div/div[2]/div[2]//button')))
            YesButton.click()
            print('YesButton.click()')
            time.sleep(8)
            counts.append(i)
            print(x6[i], "Enabled")
            i = i+1

        else:
            i = i+1
            continue
    return counts

