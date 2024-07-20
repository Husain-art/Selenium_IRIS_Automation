Content = []
NoOfContent = input("Enter No. Of Content")
for i in range(0, int(NoOfContent)):
    NoOfSL = input("Enter No. of SL for Content " + str(i+1))
    SL = []
    for j in range(0, int(NoOfSL)):
        SLCName = "|||C"+str(i+1)+"|SL"+str(j+1)
        Partition = input("Enter Percentage for " + "|||C"+str(i+1)+"|SL"+str(j+1))
        SL.append(int(Partition))
    Content.append(SL)

ContentPercentage = []
for i in range(0, len(Content)):
    Partition = input("Enter Percentage for Content" + str(i + 1))
    ContentPercentage.append(int(Partition))

filePath = "C:\\Users\\v-hcyclewala\\OneDrive - Microsoft\\V Desktop\\PythonAutomation\\Automatic_Script_Generator\\VictorCriteria.txt"
file = open(filePath, "w")
file.write("\n")
file.close()

import TemplateFile as tf

statement = tf.Numbering("MasterList", "MasterlistSplit", "Decile")
file = open(filePath, "a")
file.write(statement)
file.close()

C_LowerLimit = 0
C_UpperLimit = ContentPercentage[0]

for i in range(0, len(Content)):
    statement = tf.Content_Spliting("MasterlistSplit", "Content"+str(i+1), "Decile", str(C_LowerLimit), str(C_UpperLimit))
    file = open(filePath, "a")
    file.write(statement)
    file.close()

    statement = tf.Counts("Content"+str(i+1), "counts", "Counts for content"+str(i+1))
    file = open(filePath, "a")
    file.write(statement)
    file.close()

    statement = tf.Numbering ("Content"+str(i+1), "SegmentNo"+str(i+1), "Decile")
    file = open(filePath, "a")
    file.write(statement)
    file.close()

    S_LowerLimit = 0
    S_UpperLimit = Content[i][0]
    for j in range(0, len(Content[i])):
        if(j==0):
            statement = tf.Seg_Splitting("SegmentNo"+str(i+1), "SegmentSplit" + str(i+1))
            file = open(filePath, "a")
            file.write(statement)
            file.close()

        file = open(filePath, "r")
        List = file.readlines()
        file.close()
        LastLine = List[-1]

        Label = "|||C" + str(i + 1) + "|SL" + str(j + 1)
        Statement = tf.Condition("Decile", Label, str(S_LowerLimit), str(S_UpperLimit))
        LastLine = LastLine.replace("NULL", Statement)

        file = open(filePath, "r")
        script = file.read()
        file.close()
        script = script.replace(List[-1], LastLine)

        file = open(filePath, "w")
        file.write(script)
        file.close()

        S_LowerLimit = S_UpperLimit
        if(j != len(Content[i])-1):
            S_UpperLimit = S_UpperLimit + Content[i][j+1]

    statement = tf.Counts("SegmentSplit" + str(i+1), "counts", "Counts of users in SegmentSlit" + str(i+1))
    file = open(filePath, "a")
    file.write(statement)
    file.close()

    file = open(filePath, "a")
    file.write("\n")
    file.close()

    C_LowerLimit = C_UpperLimit
    if(i!= len(Content)-1):
        C_UpperLimit = C_UpperLimit + ContentPercentage[i+1]

# MasterList = SELECT * FROM SegmentSplit1 UNION SELECT * FROM SegmentSplit2;

file = open(filePath, "a")
file.write("MasterList = ")
file.close()

for i in range(len(Content)):
    if (i == len(Content) - 1):
        file = open(filePath, "a")
        file.write("SELECT * FROM SegmentSplit" + str(i+1) + ";")
        file.close()
    else:
        file = open(filePath, "a")
        file.write("SELECT * FROM SegmentSplit" + str(i+1) + " UNION ")
        file.close()