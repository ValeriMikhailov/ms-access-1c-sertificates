Attribute VB_Name = "mdlImportExcelTable"
Option Compare Database
Option Explicit

Function ImportExcelTable(dbPathAndName As String, funImportExcelTable As Variant)
    Dim accApp
    Dim myFunc As Variant
    Set accApp = New Access.Application
    
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    myFunc = funImportExcelTable
    
    Set myFunc = Nothing
    accApp.Quit
End Function
