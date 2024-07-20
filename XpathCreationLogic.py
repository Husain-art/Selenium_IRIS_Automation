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

ADOId = input("Enter ADOID")

website = 'https://dev.azure.com/xboxgames/Marketing_Ops/_workitems/edit/' + ADOId
service = Service(executable_path=r"C:\Users\v-hcyclewala\Downloads\edgedriver_win64\msedgedriver.exe")
driver = webdriver.Edge(service=service)
# driver = webdriver.Chrome(service = service)
driver.get(website)
driver.maximize_window()
time.sleep(8)

CampaignId = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//span[@aria-label = "ID Field"]'))).text
print("CampaignId", CampaignId)

title = driver.title
print("title", title)

PlanningElements = driver.find_element(by='xpath', value="//label[text()='Campaign Type']//parent::div//parent::div//parent::div")
x = True
Position = 1
VssAditionCount = 0
Feilds = []
CurrentRelationalValue = []

while(x):
    PlanningElementsXPath = "./div[" + str(Position) + "]//label"
    PlanningElement = PlanningElements.find_element(by='xpath', value=PlanningElementsXPath).text
    # print(PlanningElement)

    if(Position == 1):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
        continue
    else:
        PrevPlanningElementXpath = "./div[" + str(Position - 1) + "]//label"
        PrevPlanningElement = PlanningElements.find_element(by='xpath', value=PlanningElementsXPath).text
        # print(PrevPlanningElement)

    if(PrevPlanningElement == "T_Shirt Sizing"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Sponsor"):
        VssAditionCount = VssAditionCount + 10
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 2
    elif(PrevPlanningElement == "Request Type"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Campaign Type"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Program"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Campaign Category"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Channel"):
        VssAditionCount = VssAditionCount + 0
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Campaign Cadence"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "If Recurring"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Privacy Review Required (For Data Extract)"):
        VssAditionCount = VssAditionCount + 0
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif(PrevPlanningElement == "Campaign Frequency"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "XBAD Sizing"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Story Points"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Requested Launch Date"):
        VssAditionCount = VssAditionCount + 4
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Committed Date"):
        VssAditionCount = VssAditionCount + 4
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Actual Launch Date"):
        VssAditionCount = VssAditionCount + 4
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Revised Committed Date"):
        VssAditionCount = VssAditionCount + 4
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    elif (PrevPlanningElement == "Reason for Revised Date"):
        VssAditionCount = VssAditionCount + 3
        CurrentRelationalValue.append(VssAditionCount)
        Feilds.append(PlanningElement)
        Position = Position + 1
    else:
        print("/////////There might be Some ADO Changes/////////////")

    try:
        PlanningElementsXPath = "./div[" + str(Position) + "]//label"
        PlanningElement = PlanningElements.find_element(by='xpath', value=PlanningElementsXPath).text
    except NoSuchElementException:
        x = False

FeildsMapping = {Feilds[i]: CurrentRelationalValue[i] for i in range(len(CurrentRelationalValue))}

# print(FeildsMapping)

try:
    attributeValue = driver.find_element(by='xpath', value="//div[text()='Recurring']").get_attribute("id")
    # print(attributeValue)
except NoSuchElementException:
    attributeValue = driver.find_element(by='xpath', value="//div[text()='One-time']").get_attribute("id")
    # print(attributeValue)
CadenceValuesplit = attributeValue.split("_")
# print(CadenceValuesplit[-1])

CadenceValueNumber = int(CadenceValuesplit[-1])
# print("Campaign Cadence ValueNo.", CadenceValueNumber)
ActualCadenceValueNumber = CadenceValueNumber-1

adjustValue = FeildsMapping["Campaign Cadence"]
# print(adjustValue)
OriginValue = ActualCadenceValueNumber-adjustValue

ActualFeildsMapping = {Feilds[i]: CurrentRelationalValue[i] + OriginValue for i in range(len(CurrentRelationalValue))}
# print(ActualFeildsMapping)

# CurrentValue = []
# print(OriginValue)
# for i in range(len(CurrentRelationalValue)):
#     CurrentValue.append(OriginValue + CurrentRelationalValue[i])
# print(CurrentValue)
def XpathGenerator(ActualFeildsMapping, Feild):
    RequestTypeValue_XPath = "//div[@id=\"vss_" + str(ActualFeildsMapping[Feild]) + "\"]"
    # print(RequestTypeValue_XPath)
    return RequestTypeValue_XPath

CampaignCadenceValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping,"Campaign Cadence")).get_attribute("innerText")
print("CampaignCadenceValue:", CampaignCadenceValue)

CampaignFrequencyValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping, "Campaign Frequency")).get_attribute("innerText")
print("CampaignFrequencyValue:", CampaignFrequencyValue)

RequestedLaunchDateValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping, "Requested Launch Date")).get_attribute("innerText")
print("RequestedLaunchDateValue:", RequestedLaunchDateValue)

CommittedDateValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping, "Committed Date")).get_attribute("innerText")
print("CommittedDateValue:", CommittedDateValue)

ActualLaunchDateValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping, "Actual Launch Date")).get_attribute("innerText")
print("ActualLaunchDateValue:", ActualLaunchDateValue)

RevisedCommittedDateValue = driver.find_element(by='xpath', value=XpathGenerator(ActualFeildsMapping, "Revised Committed Date")).get_attribute("innerText")
print("RevisedCommittedDateValue:", RevisedCommittedDateValue)

ProgramIdValue_XPath = "//div[@id=\"vss_" + str(CurrentRelationalValue[-1]+OriginValue+3) + "\"]"
# print(ProgramIdValue_XPath)
ProgramIdValue = driver.find_element(by='xpath', value=ProgramIdValue_XPath).get_attribute("innerText")
print("ProgramIdValue:", ProgramIdValue)

from datetime import datetime
date_Format = '%m/%d/%Y %I:%M %p'

if(RevisedCommittedDateValue != ""):
    LaunchDate = datetime.strptime(RevisedCommittedDateValue, date_Format)

elif(CommittedDateValue != ""):
    LaunchDate = datetime.strptime(CommittedDateValue, date_Format)

elif(RequestedLaunchDateValue != ""):
    LaunchDate = datetime.strptime(RequestedLaunchDateValue, date_Format)
else:
    print("Lanuch date might be not available on ADO")

print(LaunchDate)

SegmentationTabXpath = driver.find_element(by='xpath', value="//div[@class='work-item-form-tabs']//li[3]")
SegmentationTabXpath.click()
time.sleep(2)
x = True
No_Of_Segments = []
SegId_VSS = []
Audience_Name = []
ControlValue = []
Markets = []
Count = 1
while(x):
    SegmentId = str(ADOId) + str("0") + str(Count)
    SegId_XpathText = "//div[text()=\'" + SegmentId + "\']"
    try:
        segmentAttributeValue = driver.find_element(by='xpath', value=SegId_XpathText).get_attribute("id")
        No_Of_Segments.append(SegmentId)
        SegId_TargetValuesplit = segmentAttributeValue.split("_")
        # print(SegId_TargetValuesplit[-1])
        SegId_TargetValueNumber = int(SegId_TargetValuesplit[-1])
        # print(SegId_TargetValueNumber)
        SegId_TargetValueNumber = SegId_TargetValueNumber-1
        # print(SegId_TargetValueNumber)
        SegId_VSS.append(SegId_TargetValueNumber)
        Audience_TargertValueNumber = SegId_TargetValueNumber + 6
        Audience_Xpath = '//div[@id = \"vss_' + str(Audience_TargertValueNumber) + "\"]"
        Audience = driver.find_element(by='xpath', value=Audience_Xpath).get_attribute("innerText")
        print("Audience", Audience)
        Audience_Name.append(Audience)
        # Market_TargertValueNumber = SegId_TargetValueNumber + 15
        # Market_Xpath = '//div[@id = \"vss_' + str(Market_TargertValueNumber) + "\"]"
        # Market = driver.find_element(by='xpath', value=Market_Xpath).get_attribute("innerText")
        # print("Market", Market)
        # Markets.append(Market)
        ControlPercent_TargertValueNumber = SegId_TargetValueNumber + 18
        ControlPercent_Xpath = '//div[@id = \"vss_' + str(ControlPercent_TargertValueNumber) + "\"]"
        ControlPercent = driver.find_element(by='xpath', value=ControlPercent_Xpath).get_attribute("innerText")
        print("ControlPercent", ControlPercent)
        ControlValue.append(ControlPercent)
        Count = Count + 1

    except NoSuchElementException:
        break

print(SegId_VSS)
print(No_Of_Segments)
print(Audience_Name)
print(ControlValue)
