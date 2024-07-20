import openpyxl
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
from selenium.webdriver.remote import webelement
from selenium.webdriver.support.wait import WebDriverWait
# link ='https://studio.iris.microsoft.com/#/segments/SM4ULDOSJF6Y9L'
# service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
# driver = webdriver.Edge(service=service)

# pasting file from IRIS Segment
import pyautogui
import os
import time

def sendingTextToNote(sn,Segment_path):
    fullFilepath = Segment_path+sn+".txt"
    # fullFilepath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonHandson\\"+sn+".txt"
    print(fullFilepath)
    open(fullFilepath, "w")

    # path = r"C:\Users\v-hcyclewala\OneDrive - Microsoft\V Desktop\PythonHandson\DummyScriptCopy.txt"
    path = os.path.realpath(fullFilepath)
    os.startfile(path)
    time.sleep(1)
    pyautogui.hotkey('ctrl','a')
    time.sleep(1)
    print("Copied")
    pyautogui.hotkey('ctrl','v')        #paset copied text from clipboard
    time.sleep(1)
    print("Pasted")
    pyautogui.hotkey('ctrl','s')        #save the file
    time.sleep(1)
    print("Saved")
    pyautogui.hotkey('alt','f4')        #close the file and back to program
    time.sleep(1)
    print("closed")

Segment_Id = []
SegmentLink = []
Segment_Name = []
Segment_Description = []
Segment_LineOfBusiness = []
Segment_Tags = []
Tag_AudienceId = []
Tag_Campaign = []
Tag_Freq = []
Segment_Product = []
Segment_UserType = []
Segment_SourceType = []
Segment_QueryType = []
Segment_DelayPublish = []

Script_WhichVC = []
Script_GetCounts = []
Script_CollectionId = []
Script_ProgramId = []
Script_AudienceId = []
Date1 = []
Job_Status1 = []
Counts1 = []
Date2 = []
Job_Status2 = []
Counts2 = []
Date3 = []
Job_Status3 = []
Counts3 = []
def FullSegmentDetails(Segment_Link, path):
    print(Segment_Link)
    InitLink = "https://studio.iris.microsoft.com/#/segments"
    service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
    driver = webdriver.Edge(service=service)
    driver.get(InitLink)
    driver.maximize_window()
    time.sleep(3)

    title = driver.title
    print(title)
    loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
    loginButton.click()
    time.sleep(5)
    i = 0
    for link in Segment_Link:
        driver.get(link)
        driver.maximize_window()
        SegmentLink.append(link)
        title = driver.title
        print(title)

        S_ID = WebDriverWait(driver, 250).until(EC.presence_of_element_located((By.XPATH, '//span[@data-test="Segment Id"]')))
        S_ID = S_ID.text
        Segment_Id.append(S_ID)
        print(S_ID)

        S_Name = driver.find_element(by='xpath', value='//span[@data-test="Name"]').text
        Segment_Name.append(S_Name)
        print(S_Name)

        S_Description = driver.find_element(by='xpath', value='//span[@data-test="Description"]').text
        Segment_Description.append(S_Description)
        print(S_Description)

        S_LineOfBusiness = driver.find_element(by='xpath', value='//span[@data-test="Line of Business"]').text
        Segment_LineOfBusiness.append(S_LineOfBusiness)
        print(S_LineOfBusiness)

        S_Tags = driver.find_element(by='xpath', value='//span[@data-test="Segment Tags"]').text
        Tags = S_Tags.split(",")
        Tag_AudienceId.insert(i, 'No Value')
        Tag_Campaign.insert(i, 'No Value')
        Tag_Freq.insert(i, 'No Value')

        for tag in Tags:
            if (tag.__contains__("ADOA")):
                Tag_AudienceId[i] = tag
            elif (tag.__contains__("ADOC")):
                Tag_Campaign[i] = tag
            elif(tag.__contains__("ADOO") or tag.__contains__("ADOR")):
                Tag_Freq[i] = tag
            else:
                continue

        Segment_Tags.append(S_Tags)
        print(S_Tags)

        S_Product = driver.find_element(by='xpath', value='//span[@data-test="Product"]').text
        Segment_Product.append(S_Product)
        print(S_Product)

        S_UserType = driver.find_element(by='xpath', value='//span[@data-test="User Type"]').text
        Segment_UserType.append(S_UserType)
        print(S_UserType)

        S_SourceType = driver.find_element(by='xpath', value='//span[@data-test="Source Type"]').text
        Segment_SourceType.append(S_SourceType)
        print(S_SourceType)

        QueryType = driver.find_element(by='xpath', value='//span[@data-test="Query Type"]').text
        Segment_QueryType.append(QueryType)
        print(QueryType)

        S_DelayPublish = driver.find_element(by='xpath', value='//span[@data-test="Delay publish for __ hour(s) after 12 AM PDT"]').text
        Segment_DelayPublish.append(S_DelayPublish)
        print(S_DelayPublish)

        nextButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/form/div[2]/button[2]')))
        nextButton.click()

        # //div[@data-test="customQuery"]'
        TextArea = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//textarea')))
        ActionChains(driver).double_click(TextArea).perform()
        time.sleep(1)
        TextArea.send_keys(Keys.CONTROL + "A")
        time.sleep(1)
        TextArea.send_keys(Keys.CONTROL + "C")

        sendingTextToNote(S_Name, path)

        nextButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//main/div/div/div[2]/div/form/div[2]/button[3]')))
        nextButton.click()

        try:
            date1 = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[@aria-rowindex="0"]/div/div[1]'))).text
            Date1.append(date1)
        except NoSuchElementException:
            Date1.append("No value found")
        # //div[@aria-rowindex="0"]/div/div[2]

        try:
            jobStatus1 = driver.find_element(by='xpath', value='//div[@aria-rowindex="0"]/div/div[2]').text
            Job_Status1.append(jobStatus1)
        except NoSuchElementException:
            Job_Status1.append("No value found")
        # //div[@aria-rowindex="0"]/div/div[2]

        try:
            counts1 = driver.find_element(by='xpath', value='//div[@aria-rowindex="0"]/div/div[3]').text
            Counts1.append(counts1)
        except NoSuchElementException:
            Counts1.append("No value found")

        #
        # date2 = driver.find_element(by='xpath', value='//div[@aria-rowindex="1"]/div/div[1]').text
        # Date2.append(date2)
        #
        # jobStatus2 = driver.find_element(by='xpath', value='//div[@aria-rowindex="1"]/div/div[2]').text
        # Job_Status2.append(jobStatus2)
        #
        # counts2 = driver.find_element(by='xpath', value='//div[@aria-rowindex="1"]/div/div[3]').text
        # Counts2.append(counts2)
        #
        # date3 = driver.find_element(by='xpath', value='//div[@aria-rowindex="2"]/div/div[1]').text
        # Date3.append(date3)
        #
        # jobStatus3 = driver.find_element(by='xpath', value='//div[@aria-rowindex="2"]/div/div[2]').text
        # Job_Status3.append(jobStatus3)
        #
        # counts3 = driver.find_element(by='xpath', value='//div[@aria-rowindex="2"]/div/div[3]').text
        # Counts3.append(counts3)

        fullFilepath = path + S_Name + ".txt"
        with open(fullFilepath, "r") as f:
            lines = f.readlines()
            Flag1 = Flag2 = Flag3 = Flag4 = Flag5 = 0
            Script_AudienceId.insert(i, 'No Value')
            Script_ProgramId.insert(i, 'No Value')
            Script_CollectionId.insert(i, 'No Value')
            Script_GetCounts.insert(i, 'No Value')
            Script_WhichVC.insert(i, 'No Value')

            for index, line in enumerate(lines):
                if line.find('#DECLARE whichVC string = ') != -1:
                    L_whichVC = lines.index(line)
                    print("Line no.", L_whichVC)
                    values = line.split("\"")
                    whichVC = values[-2]
                    # Script_WhichVC.append(whichVC)
                    Script_WhichVC[i] = whichVC
                    # Script_AudienceId.insert(i, whichVC)
                    print(whichVC)
                    Flag1 = 1

                elif line.find('#DECLARE getCounts bool = ') != -1:
                    L_getCounts = lines.index(line)
                    print("Line no.", L_getCounts)
                    values = line.split(" ")
                    getCounts = values[-1]
                    getCounts = getCounts.strip('\n')
                    getCounts = getCounts.replace(';', '')
                    # Script_GetCounts.append(getCounts)
                    Script_GetCounts[i] = getCounts
                    print(getCounts)
                    Flag2 = 1

                elif line.find('#DECLARE CollectionId string = ') != -1:
                    L_CollectionId = lines.index(line)
                    print("Line no.", L_CollectionId)
                    values = line.split("\"")
                    CollectionId = values[-2]
                    print(CollectionId)
                    # Script_CollectionId.append(CollectionId)
                    Script_CollectionId[i] = CollectionId
                    Flag3 = 1

                elif line.find('#DECLARE ProgramId string = ') != -1:
                    L_ProgramId = lines.index(line)
                    print("Line no.", L_ProgramId)
                    values = line.split("\"")
                    ProgramId = values[-2]
                    # Script_ProgramId.append(ProgramId)
                    Script_ProgramId[i] = ProgramId
                    print(ProgramId)
                    Flag4 = 1

                elif line.find('#DECLARE AudienceId string = ') != -1:
                    L_AudienceId = lines.index(line)
                    print("Line no.", L_AudienceId)
                    values = line.split("\"")
                    AudienceId = values[-2]
                    # Script_AudienceId.append(AudienceId)
                    Script_AudienceId[i] = AudienceId
                    print(AudienceId)
                    Flag5 = 1
                elif ((Flag1 == 1) and (Flag2 == 1) and (Flag3 == 1) and (Flag4 == 1) and (Flag5 == 1)):
                    break
                elif(index==len(lines)-1):
                    break
        i = i+1
    driver.quit()

    df4 = pd.DataFrame({"Segment_Id": Segment_Id, "Segment_Link": SegmentLink, "Segment_Name": Segment_Name,
                        "Segment_Description": Segment_Description, "Segment_LineOfBusiness": Segment_LineOfBusiness,
                        "Segment_Tags": Segment_Tags, "Tag_AudienceId": Tag_AudienceId, "Tag_Campaign": Tag_Campaign,
                        "Tag_Freq": Tag_Freq, "Segment_Product": Segment_Product,
                        "Segment_UserType": Segment_UserType, "Segment_SourceType": Segment_SourceType,
                        "Segment_QueryType": Segment_QueryType, "Segment_DelayPublish": Segment_DelayPublish,
                        "Script_WhichVC": Script_WhichVC, "Script_GetCounts": Script_GetCounts,
                        "Script_CollectionId": Script_CollectionId, "Script_ProgramId": Script_ProgramId,
                        "Script_AudienceId": Script_AudienceId, "Date1": Date1, "Job_Status1": Job_Status1,
                        "Counts1": Counts1
                        # "Date2": Date2, "Job_Status2": Job_Status2, "Counts2": Counts2, "Date3": Date3,
                        # "Job_Status3": Job_Status3,
                        # "Counts3": Counts3
                        })
    return df4




