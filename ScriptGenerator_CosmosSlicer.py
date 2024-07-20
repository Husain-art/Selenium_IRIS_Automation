# /////////////////////////////////////////Prior Requirement///////////////////
# SegmentName/InteractionName XGPMAUNALOAPCEMEAMarkets
# Segment/Interaction Description: All active XGP PC subscribers in select EMEA Markets with AR off
# CampaignName: XGPMAUNALOAPCTransaction
# Campaign Description: Email informing customers about upcoming changes to their Xbox Game Pass PC subscription and their options
# TreatmentName: XGPMAUNALOAPCEMEA
# /////////////////End////////////////////
import os
import re
import time
import calendar

import pandas as pd
import datetime
from ADO_Coonectivity import ADOCampaignDetail as adoDetail
def modify(filepath, from_, to_):
    file = open(filepath, "r+")
    text = file.read()
    pattern = from_
    splitted_text = re.split(pattern, text)
    modified_text = to_.join(splitted_text)
    with open(filepath, 'w') as file:
        file.write(modified_text)

def ValidationChk(String, Limit, Value):
    x = True
    while x:
        if(re.search("[^A-Za-z_0-9-]", String) or len(String) > Limit):
            x = True
            String = input("Enter " + Value + " again")
        else:
            x = False
    return String

def ValidationChkD(String, Limit, Value):
    x = True
    while x:
        if(re.search("[^A-Za-z_\s0-9-]", String) or len(String) > Limit):
            x = True
            String = input("Enter " + Value + " again")
        else:
            x = False
    return String

def remove_special_chars_isalnum(s):
    return ''.join(letter for letter in s if letter.isalnum())

Audience_Name = []
for i in range(len(adoDetail.Audience_Name)):
    Audience_Name.append(remove_special_chars_isalnum(adoDetail.Audience_Name[i]))

print(Audience_Name)

SegmentsCounts = []
SI_Name = []
SI_Description = []
ActualTreatmentName = []

title = adoDetail.title
campaignNameList = title.split(":", 1)
Actual_CampaignName = campaignNameList[1]
Actual_CampaignName = remove_special_chars_isalnum(Actual_CampaignName)
print(Actual_CampaignName)
Choice = input("Do You want to continiue with this Name: Y/N")
if(Choice=="N"):
    Actual_CampaignName = input("Enter Campaign Name of Your Choice")

Actual_CampaignName = ValidationChk(Actual_CampaignName, 72, "Campaign Name")

Actual_CampaignDescription = input("Enter Campaign Description")
Actual_CampaignDescription = ValidationChkD(Actual_CampaignDescription, 150, "Campaign Description")

# NoOfSegment = input("Enter No. Segments")
for i in range(len(adoDetail.No_Of_Segments)):
    Counts = input("Enter Counts for Segment" + str(i+1))
    SegmentsCounts.append(int(Counts))

    Name = Audience_Name[i]
    if(Name==""):
        Name = input("Please enter Segment" + str(i+1) + "name")
        Name = ValidationChk(Name, 64, "SegmentName" + str(i + 1))
    else:
        Name = ValidationChk(Name, 64, "SegmentName" + str(i+1))
        SI_Name.append(Name)
        print("Segment" + str(i+1) + "name: " + Name)

    Description = input("Enter Segment OR Interaction Description")
    Description = Description.strip()
    Description = ValidationChkD(Description, 149, "Interaction" + str(i+1) + "Description")
    SI_Description.append(Description)

    ATreatmentName = Audience_Name[i]
    if (ATreatmentName == ""):
        ATreatmentName = input("Please enter Treatment" + str(i + 1) + "name")
    ATreatmentName = ValidationChkD(ATreatmentName, 32, "TreatmentName" + str(i+1))
    ActualTreatmentName.append(ATreatmentName)
    print("Treatment" + str(i + 1) + "name: " + ATreatmentName)


ADOId = adoDetail.ADOId
ProgramId = adoDetail.ProgramIdValue
Cadence = adoDetail.CampaignCadenceValue
LaunchDate = adoDetail.LaunchDate

if ProgramId == "":
    ProgramId = LaunchDate.year[-2:-1] + LaunchDate.month + int(ADOId)

if Cadence == "One-time":
    Freq = "O"
    EndDate = LaunchDate + datetime.timedelta(days=28)
elif Cadence == "Recurring":
    Freq = "R"
    EndDate = LaunchDate + datetime.timedelta(days=365*9)
else:
    Freq = "N"
    EndDate = LaunchDate + datetime.timedelta(days=28)

Markets = input("Enter \"M\" for Multiple Markets and Multiple Language \n \"C\" for one type of markets \n \"S\" Single Market")
if Markets == "M":
    Lang_Mrkt = "mul-MUL"
elif Markets == "C":
    Language = input("Enter Language code (like en for English")
    Lang_Mrkt = Language + "-MUL"
else:
    Lang_Mrkt = input("Enter Specific Market (Like en-US)")

Platform = input("Enter the platform type:\n \"S\" for SFMC \n \"A\" for Adobe")
if Platform == "S":
    PlatformType = "SFMC"
    InstanceName = "SFMC_DEVICES_STUDIOS_XBOX_US_PROD"
elif Platform == "A":
    PlatformType = "Adobe"
    InstanceName = "Mscom-mkt-Prod7"
else:
    PlatformType = "Not Define"
    InstanceName = "Not Define"

if adoDetail.title.__contains__("Transactional"):
    TopicId = "Transactional"
else:
    TopicId = "Xbox Promotional Communications"

CampFreq = adoDetail.CampaignFrequencyValue
if CampFreq == "Daily":
    Interval = 365
else:
    Interval = 1

currentTime = datetime.datetime.now()
LastTime = datetime.datetime(2025, 5, 10, 13, 30, 00)
if currentTime.time() >= LastTime.time():
    startDate = datetime.datetime.today() + datetime.timedelta(days=1)
else:
    startDate = datetime.datetime.today()

d = datetime.datetime.strptime("14:00:00", "%H:%M:%S")
StartTime = d.strftime("%I:%M %p")

scheduleType = "Monthly"

SemiFileDropDate = LaunchDate + datetime.timedelta(days=-2)
SemiFileDropDateDay = calendar.weekday(SemiFileDropDate.year, SemiFileDropDate.month, SemiFileDropDate.day)

if SemiFileDropDateDay == 5:
    FileDropDateDay = (LaunchDate + datetime.timedelta(days=-3)).day
else:
    FileDropDateDay = (LaunchDate + datetime.timedelta(days=-2)).day

Partition = 6000000

InputfilePath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\Automatic_Script_Generator\\PreDefineScript\\CosmosSlicerTemplate.txt"
file = open(InputfilePath, "r")
slice = file.read()
file.close()

for i in range(len(SegmentsCounts)):
    LowerLimit = 0
    UpperLimit = Partition
    counts = 1
    FileName = "Segment"+str(i+1)

    OutputfilePath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\" + ADOId + "\\Automatic_Script_Generator\\CosmosSlicer"
    os.makedirs(OutputfilePath, 777, True)

    Outputfile = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\" + ADOId + "\\Automatic_Script_Generator\\CosmosSlicer\\"+ FileName +".txt"
    OPfile = open(Outputfile, "w")

    OPfile.write("\n")
    OPfile.write("MasterList = SELECT *.Except(Xuid) FROM MasterList; \n Fd= SELECT * ,ROW_NUMBER() OVER(ORDER BY Puid)AS Rr FROM MasterList;")
    OPfile.close()

    while((SegmentsCounts[i] + Partition)>UpperLimit):
        OPfile = open(Outputfile, "a")
        OPfile.write(slice)
        OPfile.close()

        modify(Outputfile, "File_No", "File_"+str(counts))
        modify(Outputfile, "LowerLimit", str(LowerLimit))
        modify(Outputfile, "UpperLimit", str(UpperLimit))
        modify(Outputfile, "FolderName", "ADO"+ADOId)
        modify(Outputfile, "FileName", "Segment"+str(i+1)+chr(counts+64))
        modify(Outputfile, "Display_Query", "Counts File_" + str(counts))
        counts = counts+1
        LowerLimit = UpperLimit
        UpperLimit = UpperLimit + Partition

S_Action = []
S_SegmentName = []
S_Product = []
S_Description = []
S_Delay_In_Publish = []
S_CampTag = []
S_AudTag = []
S_FreqTag = []
S_UserType = []
S_SegmentLink = []
S_ScriptName = []

C_CampaignName = []
C_LineOfBusiness = []
C_Description = []
C_CampaignOwningProduct = []
C_LifeCycle = []

I_InteractionName = []
I_InteractionDescription = []
I_UserType = []
I_SegmentName = []
I_Control = []
I_TreatmentName = []
I_Action = []
I_PlatformType = []
I_InstanceName = []
I_TopicId = []
I_Interval = []
I_StartDate = []
I_StartTime = []
I_scheduleType = []
I_EndDate = []
I_DayOfMonth = []
I_AddInteractionLink = []
I_Task = []

# ///////////////////////////IRIS Slicing Script////////////////////////////////////
InputfilePath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\Automatic_Script_Generator\\PreDefineScript\\IRIS_Slicer.txt"
file = open(InputfilePath, "r")
slice = file.read()
file.close()

for i in range(len(SegmentsCounts)):
    LowerLimit = 0
    UpperLimit = Partition
    counts = 1


    while((SegmentsCounts[i] + Partition)>UpperLimit):
        Sequence = str(i + 1) + chr(counts+64)
        FileName = "Segment" + Sequence

        OutputfilePath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\" + ADOId + "\\Automatic_Script_Generator\\IRISSlicer"
        os.makedirs(OutputfilePath, 777, True)

        Outputfile = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\" + ADOId + "\\Automatic_Script_Generator\\IRISSlicer\\" + FileName + ".txt"
        OPfile = open(Outputfile, "w")

        OPfile.write("\n")
        OPfile.close()
        OPfile = open(Outputfile, "a")
        OPfile.write(slice)
        OPfile.close()

        modify(Outputfile, "CollectionIdValue", ADOId)
        modify(Outputfile, "ProgramIdValue", ProgramId)
        modify(Outputfile, "AudienceIdValue", ADOId+"0"+str(i+1))
        modify(Outputfile, "FolderName", "ADO"+ADOId)
        modify(Outputfile, "FileName", FileName)

        Camp_Tag = "ADOC" + ADOId
        S_CampTag.append(Camp_Tag)
        Aud_Tag = "ADOA" + ADOId + "0" + str(i+1)
        S_AudTag.append(Aud_Tag)
        Freq_Tag = "ADO" + Freq + ProgramId
        S_FreqTag.append(Freq_Tag)
        Segment_Name = Freq_Tag + '_' + Aud_Tag + '_' + SI_Name[i] + '_Segment' + Sequence
        S_SegmentName.append(Segment_Name)
        S_Action.append('Create')
        S_Product.append('Xbox')
        S_Description.append(SI_Description[i] + " Part" + str(counts))
        S_Delay_In_Publish.append('0.5')
        S_UserType.append('PUIDINT')
        S_SegmentLink.append(' ')
        S_ScriptName.append(FileName)

        CampaignName = Camp_Tag + "_" + Freq_Tag + "_" + Actual_CampaignName + "_EM"
        C_CampaignName.append(CampaignName)
        C_LineOfBusiness.append("Consumer")
        C_Description.append(Actual_CampaignDescription)
        C_CampaignOwningProduct.append("Xbox")
        C_LifeCycle.append("")

        Interaction_Name = Freq_Tag + '_' + Aud_Tag + '_' + SI_Name[i] + '_Interaction' + Sequence
        I_InteractionName.append(Interaction_Name)
        I_InteractionDescription.append(SI_Description[i] + " Part" + str(counts))
        I_UserType.append("PUIDINT")
        I_SegmentName.append(Segment_Name)
        I_Control.append(adoDetail.ControlValue[i])
        TreatmentName = ProgramId + "_" + ADOId + "0" + str(i+1) + "_" + Lang_Mrkt + "_" + ActualTreatmentName[i] + "_T"
        I_TreatmentName.append(TreatmentName)
        I_Action.append("Email")
        I_PlatformType.append(PlatformType)
        I_InstanceName.append(InstanceName)
        I_TopicId.append(TopicId)
        I_Interval.append(Interval)
        I_StartDate.append(startDate)
        I_StartTime.append(StartTime)
        I_scheduleType.append(scheduleType)
        I_EndDate.append(EndDate)
        I_DayOfMonth.append(FileDropDateDay)
        I_AddInteractionLink.append("")
        I_Task.append("Create")

        counts = counts+1
        LowerLimit = UpperLimit
        UpperLimit = UpperLimit + Partition

Filepath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\" + ADOId + "\\IRIS_Automation\\CreateSegments\\"
os.makedirs(Filepath, 777, True)
OutputPath = Filepath + "IRISCreation" + ".xlsx"

df0 = pd.DataFrame({"Action": S_Action, "SegmentName": S_SegmentName, "Product": S_Product, "Description": S_Description, "Delay_In_Publish": S_Delay_In_Publish,  "CampTag": S_CampTag, "AudTag": S_AudTag, "FreqTag": S_FreqTag, "UserType": S_UserType, "SegmentLink": S_SegmentLink, "ScriptNaming": S_ScriptName})
df1 = pd.DataFrame({"CampaignName": C_CampaignName, "LineOfBusiness": C_LineOfBusiness, "Description": C_Description, "CampaignOwningProduct": C_CampaignOwningProduct, "LifeCycle": C_LifeCycle, "InteractionName": I_InteractionName, "InteractionDescription": I_InteractionDescription, "UserType": I_UserType,
                    "SegmentName": I_SegmentName, "Control": I_Control, "TreatmentName": I_TreatmentName, "Action": I_Action, "PlatformType": I_PlatformType, "InstanceName": I_InstanceName,
                    "TopicId": I_TopicId, "Interval": I_Interval, "StartDate": I_StartDate, "StartTime": I_StartTime, "scheduleType": I_scheduleType,
                    "EndDate": I_EndDate, "DayOfMonth": I_DayOfMonth, "AddInteractionLink": I_AddInteractionLink, "Task": I_Task})

SheetDetails = {"Segment_Detail": df0, "Campaign_Details": df1}

with pd.ExcelWriter(OutputPath) as writer:
    for sheet_Name in SheetDetails.keys():
        SheetDetails[sheet_Name].to_excel(writer, sheet_name=sheet_Name, index=False)