Attribute VB_Name = "mdlExportTableInToExcelFile"
Option Compare Database
Option Explicit

Function ExportTableInToExcelFile(tableNameForExport As String, pathToExcelFile As String)
    DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel8, tableName:=tableNameForExport, fileName:=pathToExcelFile
End Function
