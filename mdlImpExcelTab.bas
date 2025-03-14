Attribute VB_Name = "mdlImpExcelTab"
Option Compare Database
Option Explicit

Function impExcelTab(dbPathAndName As String, tabName As String, fileExcel As String)
    Dim accApp
    Set accApp = New Access.Application
    
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    accApp.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tableNameAction, pathToBase & fileExcelImport, True
    
    accApp.Quit
End Function
