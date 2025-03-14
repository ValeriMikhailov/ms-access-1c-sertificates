Attribute VB_Name = "mdlSizeBrutto"
Option Compare Database
Option Explicit

Function crtSizeDescr()
'                "[size].[" & fldSizeLength & "] * 10 AS Дл_см_мм, [size].[" & fldSizeWidth & "] * 10 AS Шир_см_мм, [size].[" & fldSizeHeight & "] * 10 AS Выс_см_мм "
    sizeDescrSQL = "SELECT [" & tabNameDescr & "].[" & fldDescrCod & "], [" & tblNameSize & "].[" & fldSizeName & "], " & _
                "[" & tblNameSize & "].[" & fldSizeWidth & "], [" & tblNameSize & "].[" & fldSizeLength & "], [" & tblNameSize & "].[" & fldSizeHeight & "], [" & tblNameSize & "].[" & fldSizeUnitLg & "], " & _
                    "[" & tblNameSize & "].[" & fldSizeWeight & "], [" & tblNameSize & "].[" & fldSizeUnitWt & "], [" & tblNameSize & "].[" & fldSizeLineNumber & "], " & _
                        "[" & tblNameSize & "].[" & fldSizeNet & "], [" & tblNameSize & "].[" & fldSizeunitNet & "],  " & _
                            "[" & tabNameDescr & "].[" & fldDescrDesc & "] " & _
                "FROM [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] ON [" & tabNameDescr & "].[" & fldDescrCod & "] = [" & tblNameSize & "].[" & fldSizeCod & "] " & _
                "ORDER BY [" & tblNameSize & "].[" & fldSizeName & "];"
    Call crtQuery(sizeDescrSQL, qrySizeDescr)
    DoCmd.OpenQuery qrySizeDescr, acViewNormal
End Function
Function crtSizeBrutto()
'1 2 3 4 5 6 7 8 9 10 11 12 13
'A B C D F G H J K M  N  P  Q
        sizeBruttoSQL = fncBruttoSQL
        Call crtQuery(sizeBruttoSQL, qrySizeBrutto)
        DoCmd.OpenQuery qrySizeBrutto, acViewNormal
End Function
Function crtCountPlaceis()
    countPlaceisSQL = "SELECT [" & tblNameSize & "].[" & fldSizeCod & "], Count([" & tblNameSize & "].[" & fldSizeLineNumber & "]) AS CountOfНомерСтроки " & _
                "FROM [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] " & _
                "GROUP BY [" & tblNameSize & "].[" & fldSizeCod & "];"
    Call crtQuery(countPlaceisSQL, qryCountPlaceis)
    DoCmd.OpenQuery qryCountPlaceis, acViewNormal
End Function
Function maxPlaceis(qryName As String, c As Integer)
Dim dbs As Database
Dim qdf As QueryDef
Dim rst As Recordset
Dim ln As Integer
    Set dbs = CurrentDb
    Set qdf = dbs.QueryDefs(qryName)
    Set rst = qdf.OpenRecordset()

    maxPlaceis = DMax(rst.Fields(c).Name, qryName)
End Function

'sQry = "sizeBruttoSQL = "
' "" & Chr(32) & Chr(38) & Chr(32) & Chr(95) <=== перенос строки SQL кода
Function fncBruttoSQL()
Dim i As Integer
Dim sQry1, sQry2, sQry3, sQry4, sQry5, sQry6, sQry7, sQry8, sQry9, sQry10, sQry11 As String
' [" & tblCountPlaceis & "].[CountOfНомерСтроки] AS pl доработать, здесь обращение к запросу countPlaceis. Можно ли напрямую к базе Size ???
sQry1 = "SELECT [" & tabNameMainBase & "].[" & fldMainCod & "],[" & tabNameMainBase & "].[" & fldMainName & "],[" & tblCountPlaceis & "].[CountOfНомерСтроки] AS pl,IIF(WN>WB,'STOP','') AS STOP,A.[нетто] AS WN, A.[" & fldSizeWeight & "]"
sQry3 = "AS WB, [1д]*[1ш]*[1в]/10^9"
sQry5 = " AS V, A.[" & fldSizeLineNumber & "] AS 1, A.[" & fldSizeWeight & "] AS 1вес, A.[" & fldSizeLength & "] AS 1д, A.[" & fldSizeWidth & "] AS 1ш, A.[" & fldSizeHeight & "] AS 1в,"
' надо убрать поле содержащее 'zzz' из sQry7, теперь это лишнее
sQry7 = "'zzz' FROM ((("
sQry9 = "SELECT * FROM [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] WHERE ((([" & tblNameSize & "].[" & fldSizeLineNumber & "])=1))) AS A "
sQry11 = "INNER JOIN [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] ON A.[Код] = [" & tblNameSize & "].[" & fldSizeCod & "]) INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] ON [" & tblNameSize & "].[" & fldSizeCod & "] = [" & tabNameMainBase & "].[" & fldMainCod & "]) INNER JOIN [" & tblCountPlaceis & "] ON [" & tblNameSize & "].[" & fldSizeCod & "] = [" & tblCountPlaceis & "].[" & fldCountPlaceisCod & "] " & _
"WHERE ((([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'Поверка*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'Доставка*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'ПНР *' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'Диагностика*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'ПО *' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'Аттестация*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like 'Монтаж*'));"

sQry2 = ""
sQry4 = ""
sQry6 = ""
sQry8 = ""
sQry10 = ""
' qryCountPlaceis доработать XXX ??? XXX
For i = 2 To maxPlaceis(qryCountPlaceis, 1) ' начинается с 2 потому, что одно товарное место есть ВСЕГДА
sQry2 = sQry2 & "+IIF(t" & i & ".[" & fldSizeWeight & "] IS NULL,0,t" & i & ".[" & fldSizeWeight & "]) "
sQry4 = sQry4 & "+IIF([" & (i) & "д] IS NULL,0,[" & (i) & "д])*IIF([" & (i) & "ш] IS NULL,0,[" & (i) & "ш])*IIF([" & (i) & "в] IS NULL,0,[" & (i) & "в])/10^9"
sQry6 = sQry6 & "t" & i & ".[" & fldSizeLineNumber & "] AS " & (i) & ", t" & i & ".[" & fldSizeWeight & "] AS " & (i) & "вес, t" & i & ".[" & fldSizeLength & "] AS " & (i) & "д, t" & i & ".[" & fldSizeWidth & "] AS " & (i) & "ш, t" & i & ".[" & fldSizeHeight & "] AS " & (i) & "в,"
sQry8 = sQry8 & "("
sQry10 = sQry10 & "LEFT JOIN (SELECT * FROM [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] WHERE ((([" & tblNameSize & "].[" & fldSizeLineNumber & "])=" & (i) & "))) AS t" & i & " ON A.[Код] = t" & i & ".[Код]) "
Next i

fncBruttoSQL = sQry1 & sQry2 & sQry3 & sQry4 & sQry5 & sQry6 & sQry7 & sQry8 & sQry9 & sQry10 & sQry11
End Function
