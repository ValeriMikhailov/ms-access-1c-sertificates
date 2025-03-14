Attribute VB_Name = "mdlEditExcelExportCSV"
Option Compare Database
Option Explicit

Public Const pathToServerExportCSV As String = "\\NVLXSRV\techdocs"
Public Const fileNameCSV As String = "\docs.csv"
'формирование и редактирование единого листа РУ+ДС+СС+УТСИ для выгрузки в CSV
Function EditExcelExportCSV()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileNameExcelToCSV)
    Set oWs = oWb.Worksheet
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False

    Dim rangeSS, rangeDS, rangeUT As Variant
    Dim lastRU, lastDS, lastSS, lastUT As Long
    
    Sheets(1).Activate
    lastRU = Cells(Rows.Count, 1).End(xlUp).Row
    
    Sheets(2).Activate
    lastDS = Cells(Rows.Count, 1).End(xlUp).Row
    Set rangeDS = Range(Cells(2, 1), Cells(lastDS, 2))
    rangeDS.Select
    Selection.Copy
    Sheets(1).Activate
    Cells(lastRU + 1, 1).PasteSpecial
    
    Sheets(3).Activate
    lastSS = Cells(Rows.Count, 1).End(xlUp).Row
    Set rangeSS = Range(Cells(2, 1), Cells(lastSS, 2))
    rangeSS.Select
    Selection.Copy
    Sheets(1).Activate
    lastRU = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lastRU + 1, 1).PasteSpecial

    Sheets(4).Activate
    lastUT = Cells(Rows.Count, 1).End(xlUp).Row
    Set rangeUT = Range(Cells(2, 1), Cells(lastUT, 2))
    rangeUT.Select
    Selection.Copy
    Sheets(1).Activate
    lastRU = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lastRU + 1, 1).PasteSpecial
    
    Call Del_SubStr
    ActiveSheet.Cells.Replace "\Doc\Part1", ""
    ActiveSheet.Cells.Replace Chr(35), ""
    ActiveSheet.Cells.Replace Chr(92) & Chr(92), Chr(47)
    ActiveSheet.Cells.Replace Chr(92), Chr(47)
    
    ActiveSheet.Rows("1:1").Delete Shift:=xlUp
'    delRows ("1:1")
    
'разделитель ";" и всё чисто
'.SaveAs fileName:="\\NVLXSRV\techdocs\docs.csv", FileFormat:=xlCSV, Local:=True, CreateBackup:=False
'    ActiveSheet.SaveAs fileName:=pathToServerExportCSV & fileNameCSV, FileFormat:=xlCSV, Local:=True, CreateBackup:=False
    
    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function
Function Del_SubStr()   'удаление строк с указанным текстом
    Dim sSubStr As String 'искомое слово или фраза(может быть указанием на ячейку)
    Dim lCol As Long 'номер столбца с просматриваемыми значениями
    Dim lLastRow As Long, li As Long
    Dim arr
 
    'Указанный текст
    sSubStr = "/mnt#\\NV3C\Doc\Part1\4_Технич.отд_ОТГРУЗ\Сертификаты и рег.удостоверения\inf.pdf#"

    'Укажите номер столбца, в котором искать указанное значение
    lCol = 1

    If lCol = 0 Then Exit Function
 
    lLastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    arr = Cells(1, lCol).Resize(lLastRow).Value
 
'    Application.ScreenUpdating = 0
    Dim rr As Range
    For li = 1 To lLastRow 'цикл с первой строки до конца
        If CStr(arr(li, 1)) = sSubStr Then
            If rr Is Nothing Then
                Set rr = Cells(li, 1)
            Else
                Set rr = Union(rr, Cells(li, 1))
            End If
        End If
    Next li
    If Not rr Is Nothing Then rr.EntireRow.Delete
'    Application.ScreenUpdating = 1
End Function
