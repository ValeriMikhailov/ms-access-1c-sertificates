Attribute VB_Name = "mdlEditExcelRegUd"
Option Compare Database
Option Explicit

Public Const shtNameRegUd = "dbRegUd"  'Лист

Function EditExcelRegUd()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathRegUdBase & fileExcelReestRegUd)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameRegUd) Then
        oXL.Sheets(shtNameRegUd).Delete
    End If

    Dim uploadRegUd As Object
    Set uploadRegUd = oXL
    
    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
    Set uploadRegUd = oXL.ActiveSheet
    uploadRegUd.Name = shtNameRegUd
 
    With oXL
'        .Range("F:H").Delete
        
'вставляем строку НАД первой строкой
'        .Range("A1").Value = "registration_number"
'        .Range("B1").Value = "registration_date"
'        .Range("C1").Value = "registration_date_end"
'        .Range("D1").Value = "name"
'        .Range("E1").Value = "producer"
'        .Range("F1").Value = "okp"
'        .Range("G1").Value = "kind"
'        .Range("C1").Value = "txtDBegin"
'        .Range("D1").Value = "txtDEnd"

'вставляем строку НАД второй строкой
        .Rows("2:2").Select
        .Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
       
        .Range("A2").Value = "a"
        .Range("C2").Value = #1/1/1900#
        .Range("F2").Value = "a"
        .Range("G2").Value = "a"
    End With

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function

