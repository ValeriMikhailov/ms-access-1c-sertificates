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
            oWs.Cells.Replace Chr(35), ""  '������� -> #
            
            Dim cell As Range, ra As Range
            Set ra = Range([A2], Range("A" & Rows.Count).End(xlUp))
            
            For Each cell In ra.Cells
                If Len(cell) Then cell.Hyperlinks.Add cell, cell    '�������������� ������ � ������ Hyperlinks
            Next cell
'�������������� ������ � ����� Excel
            If oWs.Name = "getRuExportToServer" Then oWs.Name = "��"
            If oWs.Name = "pathRegUd" Then oWs.Name = "��� ��"
            If oWs.Name = "getDS" Then oWs.Name = "��"
            If oWs.Name = "getSS" Then oWs.Name = "��"
            If oWs.Name = "pathSertificates" Then oWs.Name = "��� �� ��"
        Next
    
    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Function
