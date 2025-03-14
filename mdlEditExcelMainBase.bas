Attribute VB_Name = "mdlEditExcelMainBase"
Option Compare Database
Option Explicit

Public Const shtNameMainNomenclature = "mainNomenclature"          'Лист

Function EditExcelMainBase()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileExcelMainBase)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameMainNomenclature) Then
        oXL.Sheets(shtNameMainNomenclature).Delete
    End If

    Dim uploadMain As Object
    Set uploadMain = oXL
    
    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
    Set uploadMain = oXL.ActiveSheet
    uploadMain.Name = shtNameMainNomenclature
 
    Set oWs = oWb.Sheets(shtNameMainNomenclature)
    With oWs
        .Columns("A:N").Select
        .Selection.UnMerge
        .Range("1:4").Delete
        .Range("F:F").Delete
        .Range("B:C").Delete
        .Range("G1").Value = "поставщикКод"
        .Range("J1").Value = "txtDBgnRegU"
        .Range("K1").Value = "txtDEndRegU"
        .Range("Q1").Value = "ЕдХран"
        .Range("R1").Value = "txtDSs"
        .Range("S1").Value = "txtDel"
        .Range("T1").Value = "txtDCard"
        .Range("U1").Value = "txtDBill"
        .Range("W1").Value = "Количество"
        .Range("X1").Value = "txtTurnover"
        .Range("Y1").Value = "Артикул"
        .Range("Z1").Value = "Текстовое описание"
        .Range("AA1").Value = "ККМ"
        .Range("AB1").Value = "Автор"
        .Range("AC1").Value = "txtDDs"
        .Range("AD1").Value = "txtTrace"
        .Range("AE1").Value = "НКМИ"
        .Range("AF1").Value = "txtUTSI"
'вставляем строку НАД второй строкой
        .Rows("2:2").Select
        .Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
       
        .Range("A2").Value = "a"
        .Range("B2").Value = "a"
        .Range("E2").Value = "A a a a a a a a a a a a a a a a a A a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aa"
'        .Range("J2").Value = #1/1/3000#
        .Range("G2").Value = "a"
        .Range("N2").Value = "a"
        .Range("O2").Value = "a"
        .Range("P2").Value = "a"
'устанавливаем формат колонки General, иначе из-за разделителя Comma импортируется TEXT
        .Colums("W:W").NumberFormat = "General"
        .Range("W2").Value = 0
        .Range("Y2").Value = "a"
        .Range("Z2").Value = "A a a a a a a a a a a a a a a a a A a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aA a a a a a a a a a a a a a a a aa"
    End With

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function
