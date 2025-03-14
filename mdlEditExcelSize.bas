Attribute VB_Name = "mdlEditExcelSize"
Option Compare Database
Option Explicit

Public Const fileNameExcelSize As String = "size.xlsx"
Public Const shtNameSize = "size"          '����

Sub EditExcelSize()
    Dim oXL As Object, oWb As Object, oWs As Object
    
    On Error Resume Next
    
    Set oXL = CreateObject("Excel.Application")
    Set oWb = oXL.Workbooks.Open(pathMainBase & fileNameExcelSize)
    oXL.Visible = True
    
    oXL.Workbooks.Application.DisplayAlerts = False
    
    If oXL.SheetExists(oXL.ActiveWorkbook.Name, shtNameSize) Then
        oXL.Sheets(shtNameSize).Delete
    End If

    Dim uploadSize As Object
    Set uploadSize = oXL
    
    oXL.Sheets(1).Copy Before:=oXL.Sheets(oXL.Sheets.Count)
    Set uploadSize = oXL.ActiveSheet
    uploadSize.Name = shtNameSize
 
    With oXL
        .Columns("A:N").Select
        .Selection.UnMerge
        
        .Range("1:4").Delete
        .Range("E:E").Delete
        .Range("B:C").Delete
        .Range("C1").Value = "������"
        .Range("D1").Value = "�����"
        .Range("E1").Value = "������"
        .Range("F1").Value = "�����������"
        .Range("G1").Value = "���"
        .Range("H1").Value = "�����������"
        .Range("I1").Value = "�����������"
        .Range("J1").Value = "�����"
        .Range("K1").Value = "��������"
'��������� ������ ��� ������ �������
        .Rows("2:2").Select
        .Selection.Insert CopyOrigin:=xlFormatFromRightOrBelow
       
        .Range("A2").Value = "a"
        .Range("C2").Value = 417
        .Range("D2").Value = 417
        .Range("E2").Value = 417
        .Range("G2").Value = 0.47
        .Range("I2").Value = 417
        .Range("J2").Value = 0.417
    End With

    oWb.Save
    oXL.Quit
    Set oXL = Nothing
    Set oWb = Nothing
    Set oWs = Nothing
End Sub
