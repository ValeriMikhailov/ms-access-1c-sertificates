Attribute VB_Name = "mdlVarGeneral"
Option Compare Database
Option Explicit

'    Public accApp
'    Public excelApp
    Public pathToBase As String
    Public fileNameBase As String
    Public fileExcelImport As String
    Public tableNameAction As String
    Public nameExcelMacro As String
    Public delThisFile As String
    
'    Public dbs As Database
'    Public qdf As QueryDef
'    Public rst As Recordset

    Public Const pathMainBase As String = "D:\Work\"

' ------------ export to server excel with links ----------------
    Public Const pathToServer As String = "\\NV3C\Doc\Part1\4_������.���_������\����������� � ���.�������������\"
'    Public Const pathToServer As String = "D:\Work\TEMP\"
    
    Public Const fileInfo As String = "inf.pdf"
    Public Const fileNameExcelOnServer As String = "�� �� �� �����.xls"
    
    Public Const pathToServerCSV As String = "\\NVLXSRV\techdocs\"
'    Public Const nameExcelMacroExportToCSV As String = "testToCSV"
    Public Const nameExcelMacroExportToCSV As String = "exportToCSV"
    Public Const fileNameExcelToCSV As String = "docs.xls"

' ------------ diferentes ----------------
    Public Const frmnameFolders As String = "folders"
    Public Const frmnameSeller As String = "seller"

