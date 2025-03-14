Attribute VB_Name = "mdlExportToExcel"
Option Compare Database
Option Explicit

Function exportToExcel(dbPathAndName As String)             'экспорт таблицы в Excel
    Dim accApp
    Set accApp = New Access.Application
    
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    accApp.DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel8, tableName:=linksTableName, fileName:=pathMainBase & exportFileName
    
    Set accApp = Nothing
End Function
