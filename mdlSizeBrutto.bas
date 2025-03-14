Attribute VB_Name = "mdlSizeBrutto"
Option Compare Database
Option Explicit

Function crtSizeDescr()
'                "[size].[" & fldSizeLength & "] * 10 AS ��_��_��, [size].[" & fldSizeWidth & "] * 10 AS ���_��_��, [size].[" & fldSizeHeight & "] * 10 AS ���_��_�� "
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
    countPlaceisSQL = "SELECT [" & tblNameSize & "].[" & fldSizeCod & "], Count([" & tblNameSize & "].[" & fldSizeLineNumber & "]) AS CountOf����������� " & _
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
' "" & Chr(32) & Chr(38) & Chr(32) & Chr(95) <=== ������� ������ SQL ����
Function fncBruttoSQL()
Dim i As Integer
Dim sQry1, sQry2, sQry3, sQry4, sQry5, sQry6, sQry7, sQry8, sQry9, sQry10, sQry11 As String
' [" & tblCountPlaceis & "].[CountOf�����������] AS pl ����������, ����� ��������� � ������� countPlaceis. ����� �� �������� � ���� Size ???
sQry1 = "SELECT [" & tabNameMainBase & "].[" & fldMainCod & "],[" & tabNameMainBase & "].[" & fldMainName & "],[" & tblCountPlaceis & "].[CountOf�����������] AS pl,IIF(WN>WB,'STOP','') AS STOP,A.[�����] AS WN, A.[" & fldSizeWeight & "]"
sQry3 = "AS WB, [1�]*[1�]*[1�]/10^9"
sQry5 = " AS V, A.[" & fldSizeLineNumber & "] AS 1, A.[" & fldSizeWeight & "] AS 1���, A.[" & fldSizeLength & "] AS 1�, A.[" & fldSizeWidth & "] AS 1�, A.[" & fldSizeHeight & "] AS 1�,"
' ���� ������ ���� ���������� 'zzz' �� sQry7, ������ ��� ������
sQry7 = "'zzz' FROM ((("
sQry9 = "SELECT * FROM [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] WHERE ((([" & tblNameSize & "].[" & fldSizeLineNumber & "])=1))) AS A "
sQry11 = "INNER JOIN [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] ON A.[���] = [" & tblNameSize & "].[" & fldSizeCod & "]) INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] ON [" & tblNameSize & "].[" & fldSizeCod & "] = [" & tabNameMainBase & "].[" & fldMainCod & "]) INNER JOIN [" & tblCountPlaceis & "] ON [" & tblNameSize & "].[" & fldSizeCod & "] = [" & tblCountPlaceis & "].[" & fldCountPlaceisCod & "] " & _
"WHERE ((([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '�������*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '��������*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '��� *' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '�����������*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '�� *' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '����������*' And ([" & tabNameMainBase & "].[" & fldMainName & "]) Not Like '������*'));"

sQry2 = ""
sQry4 = ""
sQry6 = ""
sQry8 = ""
sQry10 = ""
' qryCountPlaceis ���������� XXX ??? XXX
For i = 2 To maxPlaceis(qryCountPlaceis, 1) ' ���������� � 2 ������, ��� ���� �������� ����� ���� ������
sQry2 = sQry2 & "+IIF(t" & i & ".[" & fldSizeWeight & "] IS NULL,0,t" & i & ".[" & fldSizeWeight & "]) "
sQry4 = sQry4 & "+IIF([" & (i) & "�] IS NULL,0,[" & (i) & "�])*IIF([" & (i) & "�] IS NULL,0,[" & (i) & "�])*IIF([" & (i) & "�] IS NULL,0,[" & (i) & "�])/10^9"
sQry6 = sQry6 & "t" & i & ".[" & fldSizeLineNumber & "] AS " & (i) & ", t" & i & ".[" & fldSizeWeight & "] AS " & (i) & "���, t" & i & ".[" & fldSizeLength & "] AS " & (i) & "�, t" & i & ".[" & fldSizeWidth & "] AS " & (i) & "�, t" & i & ".[" & fldSizeHeight & "] AS " & (i) & "�,"
sQry8 = sQry8 & "("
sQry10 = sQry10 & "LEFT JOIN (SELECT * FROM [" & pathMainBase & fileNameSize & "].[" & tblNameSize & "] WHERE ((([" & tblNameSize & "].[" & fldSizeLineNumber & "])=" & (i) & "))) AS t" & i & " ON A.[���] = t" & i & ".[���]) "
Next i

fncBruttoSQL = sQry1 & sQry2 & sQry3 & sQry4 & sQry5 & sQry6 & sQry7 & sQry8 & sQry9 & sQry10 & sQry11
End Function
