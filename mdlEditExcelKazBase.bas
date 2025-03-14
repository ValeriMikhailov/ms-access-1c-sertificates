Attribute VB_Name = "mdlEditExcelKazBase"
Option Compare Database
Option Explicit

Public Const fileNameExcelKazBase As String = "kaz.xlsx"
Public Const shtNameKazBase = "kazBase"          'Лист

Function EditExcelKazBase()
'    SetOption ("Confirm Action Queries"), 0
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileNameExcelKazBase)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameKazBase) Then
        oXL.Sheets(shtNameKazBase).Delete
    End If

    Dim uploadKazBase As Object
    Set uploadKazBase = oXL
    
    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
    Set uploadKazBase = oXL.ActiveSheet
    uploadKazBase.Name = shtNameKazBase
    
    Set oWs = oWb.Sheets(shtNameKazBase)
    With oWs
'        .Columns("J:K").Delete
'        .Columns("F:F").Delete
        .Columns("A:A").Delete
    End With
    
    Range("A1").Value = "cod"
    Range("B1").Value = "articule"
    Range("C1").Value = "wName"
    Range("D1").Value = "pName"
    Range("E1").Value = "unit"
    Range("G1").Value = "unit_st"
    Range("H1").Value = "price"
    Range("I1").Value = "currency"
    Range("L1").Value = "NDS"
    Range("M1").Value = "descrip"
    Range("O1").Value = "itemType"
    Range("P1").Value = "author"
    Range("Q1").Value = "textDate"
    Range("R1").Value = "groupID"
    Range("S1").Value = "grName"
   
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown

    Range("A2").Value = "a"
    Range("B2").Value = "a"
    Range("D2").Value = "A a a a a a a a a a a a a a a a a A a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aa"
    Range("L2").Value = "a"

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function
