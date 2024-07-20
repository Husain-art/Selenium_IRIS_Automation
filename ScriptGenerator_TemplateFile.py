def Numbering(Input, Output, Variable):
    Statement  = Output + " = SELECT *, NTILE(100) OVER(ORDER BY Guid.NewGuid()) AS " + Variable + " FROM " + Input + ";\n"
    return Statement

def Content_Spliting(Input, Output, Variable, LowerLimit, UpperLimit):
    Statement = Output + " = SELECT *.Except("+Variable+") FROM " + Input +" AS A WHERE " + Variable + " > " + LowerLimit +" AND "+ Variable + " <= " + UpperLimit + ';\n'
    return Statement

def Seg_Splitting(Input, Output):
    Statement = Output + " = SELECT A.*, NULL AS SubjectLine FROM " + Input + " AS A ;\n"
    return Statement

def Condition(Variable, Label, LowerLimit, UpperLimit):
    Statement = "(" + Variable + " > " + LowerLimit + " AND " + Variable + " <= " + UpperLimit + " ? \"" + Label + "\" : NULL)"
    return Statement

def Counts(Input, Output, FilterName):
    Statement = Output + " = SELECT * FROM " + Output + " UNION SELECT \"" + FilterName + "\" AS QueryRow, COUNT(DISTINCT Puid) AS  dis_CountOfPuids, COUNT(Puid) AS CountOfPuids FROM " + Input +";\n"
    return Statement