from selenium import webdriver
from selenium.webdriver import Keys
# from selenium .webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
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
# pasting file from IRIS Segment
CampainLink = "https://studio.iris.microsoft.com/#/campaigns/5052223"
Filepath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\"
website = (CampainLink)
service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
driver = webdriver.Edge(service=service)
driver.get(website)
driver.maximize_window()
time.sleep(1)
# title = driver.title
# print(title)

time.sleep(5)
loginButton = driver.find_element(by='xpath', value='//button[@class = "ms-Button ms-Button--primary root-83"]')
loginButton.click()
time.sleep(2)

C_Name = []
C_Description = []
C_Line_of_Business = []
C_Product = []
C_Lifecycle = []
C_State = []
C_Manager = []
C_Last_Updated = []

Campaign_Name = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//span[@data-test="Campaign Name"]'))).text
C_Name.append(Campaign_Name)
print(Campaign_Name)

# Description = driver.find_element(by='xpath', value='//span[@data-test="Description"]').text
Description = driver.find_element(By.XPATH, '//span[@data-test="Description"]').text
C_Description.append(Description)
print(Description)

Line_of_Business = driver.find_element(by='xpath', value='//span[@data-test="Line of Business"]').text
C_Line_of_Business.append(Line_of_Business)
print(Line_of_Business)

Product = driver.find_element(by='xpath', value='//span[@data-test="Product"]').text
C_Product.append(Product)
print(Product)

Lifecycle = driver.find_element(by='xpath', value='//span[@data-test="Lifecycle"]').text
C_Lifecycle.append(Lifecycle)
print(Lifecycle)

State = driver.find_element(by='xpath', value='//span[@data-test="State"]').text
C_State.append(State)
print(State)

Manager = driver.find_element(by='xpath', value='//span[@data-test="Manager"]').text
C_Manager.append(Manager)
print(Manager)

Last_Updated = driver.find_element(by='xpath', value='//span[@data-test="Last Updated"]').text
C_Last_Updated.append(Last_Updated)
print(Last_Updated)


InteractionsList = driver.find_elements(by='xpath', value='//div[@data-test="Interactions"]//div[(@class="ms-DetailsRow-fields fields_4b03de68") and (@data-automationid="DetailsRowFields") and (@role="presentation")]')

IntLinks = []
createdBy = []

for interaction in InteractionsList:
    interactionlink = interaction.find_element(by='xpath', value='./div[1]/a')
    IntLinks.append(interactionlink.get_attribute('href'))

    createdBy.append(interaction.find_element(by='xpath', value='./div[6]').text)

time.sleep(1)
driver.quit()

import InteractionDetails as Fi
Idf = Fi.FullInteractionDetails(IntLinks)

import SegmentDetails as FS
df = FS.FullSegmentDetails(Idf['Segment_Link'])

df0 = pd.DataFrame({"Campaign_Name": C_Name, "Description": C_Description, "Line_of_Business": C_Line_of_Business, "Product": C_Product, "Lifecycle": C_Lifecycle, "State": C_State, "Manager": C_Manager, "Last_Updated": C_Last_Updated})
df1 = pd.DataFrame({"InteractionName": Idf['InteractionName'], "InteractionLinks": Idf['InteractionLinks'], "InteractionDescription": Idf['InteractionDescription'], "Segment_Id": Idf['Segment_Id'], "Segment_Link": Idf['Segment_Link'], "Segment_Name": Idf['Segment_Name'], "State": Idf['State'], "ApprovalState": Idf['ApprovalState'], "PersonaType": Idf['PersonaType'], "InteractionState": Idf['InteractionState'], "closed": Idf['closed'], "LastUpdated": Idf['LastUpdated'], "RampPercentage": Idf['RampPercentage'], "Audience_Experiment_on": Idf['Audience_Experiment_on'], "Outcome": Idf['Outcome'], "Control": Idf['Control'], "TargetPercent": Idf['TargetPercent']})
df2 = pd.DataFrame({"InteractionName": Idf['InteractionName'], "TreatmentName": Idf['TreatmentName'], "ActionName": Idf['ActionName'], "PayloadId": Idf['PayloadId'], 	"PlatformType":Idf['PlatformType'], "Comm_Pref_Topic": Idf['Comm_Pref_Topic'], "Privacy_Comm_Type": Idf['Privacy_Comm_Type'], "Instance_Name":Idf['Instance_Name'], "DataExtension": Idf['DataExtension'], "TimeBtwView": Idf['TimeBtwView']})
df3 = pd.DataFrame({"InteractionName": Idf['InteractionName'], "StartDate": Idf['StartDate'], "EndDate": Idf['EndDate'], "ScheduleType": Idf['ScheduleType'], "ScheduleDay": Idf['ScheduleDay']})
df4 = pd.DataFrame({"Segment_Id": df['Segment_Id'], "Segment_Link": df['Segment_Link'], "Segment_Name": df['Segment_Name'], "Segment_Description": df['Segment_Description'], "Segment_LineOfBusiness": df['Segment_LineOfBusiness'], "Segment_Tags": df['Segment_Tags'], "Segment_Product": df['Segment_Product'], "Segment_UserType": df['Segment_UserType'], "Segment_SourceType": df['Segment_SourceType'], "Segment_QueryType": df['Segment_QueryType'], "Segment_DelayPublish": df['Segment_DelayPublish']})
df5 = pd.DataFrame({"Segment_Name": df['Segment_Name'], "Date1": df['Date1'], "Job_Status1": df['Job_Status1'], "Counts1": df['Counts1']})
#, "Date2": df['Date2'], "Job_Status2": df['Job_Status2'], "Counts2": df['Counts2'], "Date3": df['Date3'], "Job_Status3": df['Job_Status3'], "Counts3": df['Counts3']
SheetDetails = {"Campaign_Detail": df0, "Interaction_Detail": df1, "Treatment_Detail": df2, "Schedule_Detail": df3, "SegmentDetails": df4, "InSegmentDetails": df5}
# Segment_Id Segment_Description
# writer = pd.ExcelWriter(r"C:\Users\v-hcyclewala\OneDrive - Microsoft\V Desktop\PythonHandson\Forza.xlsx", engine="xlsxwriter")
from datetime import datetime

current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
str_current_datetime = str(current_datetime)

OutputPath = Filepath+"Output Campaign"+str_current_datetime+".xlsx"
with pd.ExcelWriter(OutputPath) as writer:
    for sheet_Name in SheetDetails.keys():
        SheetDetails[sheet_Name].to_excel(writer, sheet_name=sheet_Name, index=False)

# writer.save()
# df = pd.DataFrame({"Campaign": Campaign, "InteractionLinks": IntLinks, "Description": Description, "ApprovalState": ApprovalState, "InteractionState": InteractionState, "ScheduleType": ScheduleType, "createdBy": createdBy})
# df1.to_excel(r"C:\Users\v-hcyclewala\OneDrive - Microsoft\V Desktop\PythonHandson\PythonImport1.xlsx", index=False)
