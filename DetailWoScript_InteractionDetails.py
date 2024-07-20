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

IntLinks = []
Campaign = []
InteractionName = []
InteractionDescription = []
State = []
ApprovalState = []
PersonaType = []
InteractionState = []
closed = []
LastUpdated = []
RampPercentage = []
Audience_Experiment_on = []

Outcome = []
Control = []
TargetPercent = []

Segment_Id = []
Segment_Link = []
Segment_Name = []

TreatmentName = []
ActionName = []
PayloadId = []
PlatformType = []
Comm_Pref_Topic = []
Privacy_Comm_Type = []
Instance_Name = []
DataExtension = []
TimeBtwView = []

StartDate = []
EndDate = []
ScheduleType = []
ScheduleDay = []
def FullInteractionDetails(Interaction_Link):
    print(Interaction_Link)
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
    for interactionLink in Interaction_Link:
        driver.get(interactionLink)
        IntLinks.append(interactionLink)
        c_Name = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="0"]/div/div/div[2]/span/a')))
        c_Name = c_Name.text
        Campaign.append(c_Name)
        print(c_Name)

        I_Name = driver.find_element(by='xpath',
                                     value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="1"]/div/div/div[2]/span').text
        InteractionName.append(I_Name)
        print(I_Name)

        I_Description = driver.find_element(by='xpath',
                                            value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="2"]/div/div/div[2]/span').text
        InteractionDescription.append(I_Description)
        print(I_Description)

        I_State = driver.find_element(by='xpath',
                                      value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="3"]/div/div/div[2]/span').text
        State.append(I_State)
        print(I_State)

        I_ApprovalState = driver.find_element(by='xpath',
                                              value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="4"]/div/div/div[2]/span').text
        ApprovalState.append(I_ApprovalState)
        print(I_ApprovalState)

        I_PersonnaType = driver.find_element(by='xpath',
                                             value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="5"]/div/div/div[2]/span').text
        PersonaType.append(I_PersonnaType)
        print(I_PersonnaType)

        I_InteractionState = driver.find_element(by='xpath',
                                                 value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="6"]/div/div/div[2]/span').text
        InteractionState.append(I_InteractionState)
        print(I_InteractionState)

        I_Closed = driver.find_element(by='xpath',
                                       value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="7"]/div/div/div[2]/span').text
        closed.append(I_Closed)
        print(I_Closed)

        I_LastUpdated = driver.find_element(by='xpath',
                                            value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="8"]/div/div/div[2]/span').text
        LastUpdated.append(I_LastUpdated)
        print(I_LastUpdated)

        I_RampPercentage = driver.find_element(by='xpath',
                                               value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="9"]/div/div/div[2]/span').text
        RampPercentage.append(I_RampPercentage)
        print(I_RampPercentage)

        I_AudienceExperimentOn = driver.find_element(by='xpath',
                                                     value='//div[@data-test="Interaction summary"]//div[@class="ms-List-surface"]/div[2]/div[@data-list-index="11"]/div/div/div[2]/span').text
        Audience_Experiment_on.append(I_AudienceExperimentOn)
        print(I_AudienceExperimentOn)

        # # time.sleep(10)
        try:
            OutcomeButton = WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH,'//div[@data-test="Outcomes"]/button')))
            OutcomeButton.click()
            # # //div[@data-test="Outcomes"]//div[@aria-rowcount="2"]/div[2]//div[@data-automationid="DetailsRowFields"]/div[1]/span
            try:
                I_Outcome = driver.find_element(by='xpath', value='//div[@data-test="Outcomes"]//div[@aria-rowcount="2"]/div[2]//div[@aria-colindex="0"]/span').text
                Outcome.append(I_Outcome)
                print(I_Outcome)
            except NoSuchElementException:
                Outcome.append("No value selected")
        except NoSuchElementException:
            Outcome.append("Select Outcome section is not present")

        I_Control = driver.find_element(by='xpath', value='//div[contains(@data-test, "Control")]/button//div').text
        Control.append(I_Control)
        print(I_Control)

        AudienceButton = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@data-test, "Audience")]/button')))
        AudienceButton.click()

        S_ID = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, '//span[@data-test="Segment Id"]')))
        S_ID = S_ID.text
        Segment_Id.append(S_ID)
        print(S_ID)

        S_link = driver.find_element(by='xpath', value='//div[contains(@data-test, "Audience")]//div[@class="ms-List-surface"]/div[1]/div[@data-list-index="0"]/div/div/div[2]/span/a')
        Segment_Link.append(S_link.get_attribute('href'))
        print(S_link.get_attribute('href'))

        S_Name = driver.find_element(by='xpath', value='//span[@data-test="Name"]').text
        Segment_Name.append(S_Name)
        print(S_Name)

        T_Name = driver.find_element(by='xpath', value='//div[contains(@data-test, "Audience")]/div/div[2]').get_attribute("data-test")
        x = T_Name.split(":", 1)
        TreatmentName.append(x[0])
        TargetPercent.append(x[1])
        print(x[0])


        TreatmentButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@data-test, "Audience")]/div/div[2]/button')))
        TreatmentButton.click()

        ActionButton = WebDriverWait(driver, 420).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@data-test, "Audience")]/div/div[2]/div/div/div/button')))
        ActionButton.click()

        A_ActionName = driver.find_element(by='xpath', value='//div[contains(@data-test, "Audience")]/div/div[2]/div/div/div').get_attribute("data-test")
        ActionName.append(A_ActionName)
        print(A_ActionName)

        A_PayloadId = driver.find_element(by='xpath', value='//span[@data-test="Payload Id"]').text
        PayloadId.append(A_PayloadId)
        print(A_PayloadId)

        A_PlatformType = driver.find_element(by='xpath', value='//span[@data-test="Platform Type"]').text
        PlatformType.append(A_PlatformType)
        print(A_PlatformType)

        A_Comm_Pref_Topic = driver.find_element(by='xpath', value='//span[@data-test="Communication Preference Topic"]').text
        Comm_Pref_Topic.append(A_Comm_Pref_Topic)
        print(A_Comm_Pref_Topic)

        A_Privacy_Comm_Type = driver.find_element(by='xpath', value='//span[@data-test="Privacy Communication Type"]').text
        Privacy_Comm_Type.append(A_Privacy_Comm_Type)
        print(A_Privacy_Comm_Type)

        A_Instance_Name = driver.find_element(by='xpath', value='//span[@data-test="Instance Name"]').text
        Instance_Name.append(A_Instance_Name)
        print(A_Instance_Name)

        try:
            A_DataExtension = driver.find_element(by='xpath', value='//span[@data-test="Data Extension"]').text
            DataExtension.append(A_DataExtension)
            print(A_DataExtension)
        except NoSuchElementException:
            DataExtension.append("May be not going with SFMC")

        A_TimeBtwView = driver.find_element(by='xpath', value='//span[@data-test="Time between Views"]').text
        TimeBtwView.append(A_TimeBtwView)
        print(A_TimeBtwView)

        ScheduleButton = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//div[contains(@data-test,"Schedule")]/button')))
        ScheduleButton.click()

        Sh_StartDate = driver.find_element(by='xpath', value='//span[@data-test="Start Date"]').text
        StartDate.append(Sh_StartDate)
        print(Sh_StartDate)

        Sh_EndDate = driver.find_element(by='xpath', value='//span[@data-test="End Date"]').text
        EndDate.append(Sh_EndDate)
        print(Sh_EndDate)

        Sh_ScheduleType = driver.find_element(by='xpath', value='//span[@data-test="Schedule Type"]').text
        ScheduleType.append(Sh_ScheduleType)
        print(Sh_ScheduleType)

        Sh_ScheduleDay = driver.find_element(by='xpath', value='//span[@data-test="Schedule Interval"]').text
        ScheduleDay.append(Sh_ScheduleDay)
        print(Sh_ScheduleDay)

        time.sleep(2)

    print(Segment_Link)
    driver.quit()
    df1 = pd.DataFrame({"InteractionName": InteractionName, "InteractionLinks": IntLinks,
                        "InteractionDescription": InteractionDescription, "Segment_Id": Segment_Id,
                        "Segment_Link": Segment_Link, "Segment_Name": Segment_Name, "State": State,
                        "ApprovalState": ApprovalState, "PersonaType": PersonaType,
                        "InteractionState": InteractionState, "closed": closed, "LastUpdated": LastUpdated,
                        "RampPercentage": RampPercentage, "Audience_Experiment_on": Audience_Experiment_on,
                        "Outcome": Outcome, "Control": Control, "TargetPercent": TargetPercent,
                        "TreatmentName": TreatmentName, "ActionName": ActionName,
                        "PayloadId": PayloadId, "PlatformType": PlatformType, "Comm_Pref_Topic": Comm_Pref_Topic,
                        "Privacy_Comm_Type": Privacy_Comm_Type, "Instance_Name": Instance_Name,
                        "DataExtension": DataExtension, "TimeBtwView": TimeBtwView, "StartDate": StartDate,
                        "EndDate": EndDate, "ScheduleType": ScheduleType, "ScheduleDay": ScheduleDay})
    return df1
