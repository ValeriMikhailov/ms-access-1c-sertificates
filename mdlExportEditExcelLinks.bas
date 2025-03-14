Attribute VB_Name = "mdlExportEditExcelLinks"
Option Compare Database
Option Explicit

Function EditExcelExportLinks()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileNameExcelOnServer)
    Set oWs = oWb.Worksheet
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
        For Each oWs In Sheets
            oWs.Activate
            oWs.Cells.Replace Chr(35), ""  'изъятие -> #
            
            Dim cell As Range, ra As Range
            Set ra = Range([A2], Range("A" & Rows.Count).End(xlUp))
            
            For Each cell In ra.Cells
                If Len(cell) Then cell.Hyperlinks.Add cell, cell    'форматирование ячейки в формат Hyperlinks
            Next cell
'переименование листов в файле Excel
            If oWs.Name = "getRuExportToServer" Then oWs.Name = "РУ"
            If oWs.Name = "pathRegUd" Then oWs.Name = "все РУ"
            If oWs.Name = "getDS" Then oWs.Name = "ДС"
            If oWs.Name = "getSS" Then oWs.Name = "СС"
            If oWs.Name = "pathSertificates" Then oWs.Name = "все ДС СС"
        Next
    
    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function
