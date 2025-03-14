Attribute VB_Name = "mdlEditExcelReceipts"
Option Compare Database
Option Explicit

Function EditExcelReceipts()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileExcelReceipts)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
'    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameRegUd) Then
'        oXL.Sheets(shtNameRegUd).Delete
'    End If

'    Dim uploadReceipts As Object
'    Set uploadReceipts = oXL
    
'    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
'    Set uploadReceipts = oXL.ActiveSheet
'    uploadReceipts.Name = shtNameRegUd
 
    With oXL
        .Range("A1").Value = "txtDDate"
        .Range("B1").Value = "КодПоставщик"
        .Range("D1").Value = "Код"
'вставляем строку НАД второй строкой
        .Rows("2:2").Select
        .Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
       
        .Range("B2").Value = "a"
        .Range("D2").Value = "a"
    End With

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function


