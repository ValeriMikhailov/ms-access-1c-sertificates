Attribute VB_Name = "mdlStartExcelFileMacro"
Option Compare Database
Option Explicit

'This function starting macro of Excel Application
'pathFileExcelMacro - path to file with Excel macro
'nameExcelMacro - name of macro in Excel file
Function StartExcelFileMacro(pathFileExcelMacro As String, nameExcelMacro As String)
    SetOption ("Confirm Action Queries"), 0
    Dim objXL As Object
    
    On Error Resume Next
    Set objXL = CreateObject("Excel.Application")
    With objXL.Application
        .Visible = True
        .Workbooks.Open pathFileExcelMacro
        .Run nameExcelMacro
    End With
    
    Set objXL = Nothing
End Function
