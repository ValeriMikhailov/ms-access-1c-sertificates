Attribute VB_Name = "mdlTNVED"
Option Compare Database
Option Explicit

Function crtTNVED()
    tnvedSQL = "SELECT g.linkNumT4,g.MaxOfdateRow,[desc3] & ', ' & [desc4] AS description " & _
    "FROM (SELECT linkNumT3,desc3,MaxDateRow,desc30 " & _
        "FROM (SELECT [1colNum],[2colNum],[1colNum] & [2colNum] AS linkNumT3, Max(dateRow) AS MaxDateRow " & _
            "FROM [" & pathMainBase & "tnved.accdb].TNVED3 " & _
            "GROUP BY [1colNum] & [2colNum],[1colNum],[2colNum]) AS a " & _
                "LEFT JOIN " & _
                    "(SELECT [1colNum],[2colNum],dateRow, UCase(Left([desc30],1)) & Right([desc30],Len([desc30])-1) AS desc3, LCase([DescT3]) AS desc30 " & _
                        "FROM [" & pathMainBase & "tnved.accdb].TNVED3) AS b " & _
                        "ON (a.MaxDateRow=b.dateRow) AND (a.[1colNum]=b.[1colNum]) AND (a.[2colNum]=b.[2colNum])) AS f " & _
            "LEFT JOIN " & _
                "(SELECT linkNumT3,linkNumT4,desc4,MaxOfdateRow " & _
                    "FROM (SELECT [1colNum],[2colNum],[3colNum],LCase([descrT4]) AS desc4, dateRow " & _
                        "FROM [" & pathMainBase & "tnved.accdb].TNVED4) AS c " & _
            "RIGHT JOIN " & _
                "(SELECT [1colNum], [2colNum], [3colNum], Max(dateRow) AS MaxOfdateRow, [1colNum] & [2colNum] & [3colNum] AS linkNumT4, [1colNum] & [2colNum] AS linkNumT3, dateOut " & _
                    "FROM [" & pathMainBase & "tnved.accdb].TNVED4 " & _
                    "GROUP BY [1colNum], [2colNum], [3colNum], [1colNum] & [2colNum] & [3colNum], [1colNum] & [2colNum], dateOut " & _
            "HAVING (((dateOut) Is Null))) AS d " & _
            "ON (c.[1colNum]=d.[1colNum]) AND (c.[2colNum]=d.[2colNum]) AND (c.[3colNum]=d.[3colNum]) AND (c.dateRow=d.MaxOfdateRow)) AS g " & _
            "ON (f.linkNumT3=g.linkNumT3) WHERE (((g.linkNumT4) IS NOT NULL)) ORDER BY g.linkNumT4;"
    Call crtQuery(tnvedSQL, qryTNVED)
    DoCmd.OpenQuery qryTNVED, acViewNormal
End Function
Function crtTNVED1()
    Call crtQuery(tnvedSQL, qryTNVED1)
    DoCmd.OpenQuery qryTNVED1, acViewNormal
End Function
Function crtTnvedALL()
'    tnvedSQLALL = "SELECT g.linkNumT4,[desc3] & ', ' & [desc4] AS description,g.MaxOfdateRow,d.dateOut FROM " & _
            "(SELECT linkNumT3,desc3,MaxDateRow,desc30 FROM " & _
            "(SELECT [1colNum],[2colNum],[1colNum] & [2colNum] AS linkNumT3, Max(dateRow) AS MaxDateRow " & _
            "FROM [" & pathMainBase & "tnved.accdb].TNVED3 " & _
            "GROUP BY [1colNum] & [2colNum],[1colNum],[2colNum]) AS a " & _
            "LEFT JOIN " & _
            "(SELECT [1colNum],[2colNum],dateRow, UCase(Left([desc30],1)) & Right([desc30],Len([desc30])-1) AS desc3, LCase([DescT3]) AS desc30 " & _
            "FROM [" & pathMainBase & "tnved.accdb].TNVED3) AS b " & _
            "ON (a.MaxDateRow=b.dateRow) AND (a.[1colNum]=b.[1colNum]) AND (a.[2colNum]=b.[2colNum])) AS f " & _
            "LEFT JOIN " & _
            "(SELECT linkNumT3,linkNumT4,desc4,MaxOfdateRow,dateOut FROM " & _
            "(SELECT [1colNum],[2colNum],[3colNum],LCase([descrT4]) AS desc4,dateRow " & _
            "FROM [" & pathMainBase & "tnved.accdb].TNVED4) AS c " & _
            "RIGHT JOIN " & _
            "(SELECT [1colNum], [2colNum], [3colNum], Max(dateRow) AS MaxOfdateRow, [1colNum] & [2colNum] & [3colNum] AS linkNumT4, [1colNum] & [2colNum] AS linkNumT3, dateOut " & _
            "FROM [" & pathMainBase & "tnved.accdb].TNVED4 " & _
            "GROUP BY [1colNum], [2colNum], [3colNum], [1colNum] & [2colNum] & [3colNum], [1colNum] & [2colNum], dateOut) AS d " & _
            "ON (c.[1colNum]=d.[1colNum]) AND (c.[2colNum]=d.[2colNum]) AND (c.[3colNum]=d.[3colNum]) AND (c.dateRow=d.MaxOfdateRow)) AS g " & _
            "ON (f.linkNumT3=g.linkNumT3) WHERE (((g.linkNumT4) IS NOT NULL)) ORDER BY g.linkNumT4;"
    tnvedSQLALL = "SELECT [tnved],[descrT4],[dateRow],[dateOut] FROM [" & pathMainBase & "tnved.accdb].TNVED ORDER BY [tnved];"
    Call crtQuery(tnvedSQLALL, qryTNVED2)
    DoCmd.OpenQuery qryTNVED2, acViewNormal
End Function
