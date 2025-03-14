Attribute VB_Name = "mdlEditExcelGTD"
Option Compare Database
Option Explicit

Public Const fileNameExcelGTD As String = "GTD.xlsx"
Public Const shtNameGTD = "GTDinBase"          'Лист

Sub EditExcelGTD()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next

    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileNameExcelGTD)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameGTD) Then
        oXL.Sheets(shtNameGTD).Delete
    End If

    Dim uploadGTD As Object
    Set uploadGTD = oXL
    
    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
    Set uploadGTD = oXL.ActiveSheet
    uploadGTD.Name = shtNameGTD
 
    With oXL
        .Columns("A:F").Select
        .Selection.UnMerge
        
        .Range("1:4").Delete
        .Range("E:E").Delete
        .Range("B:C").Delete
        .Range("B1").Value = "ГТД"
'вставляем строку НАД второй строкой
        .Rows("2:2").Select
        .Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
       
        .Range("A2").Value = "a"
    End With

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Sub

