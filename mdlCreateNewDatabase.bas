Attribute VB_Name = "mdlCreateNewDatabase"
Option Compare Database
Option Explicit

Function CreateNewDatabase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    
    accApp.NewCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    
    accApp.Quit
End Function
