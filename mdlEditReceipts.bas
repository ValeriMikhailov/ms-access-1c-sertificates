Attribute VB_Name = "mdlEditReceipts"
Option Compare Database
Option Explicit

Function editReceipts(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    Call crtRegUdSQLdelRow
    
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameReceipts)
'разборка полей содержащих ДАТЫ
    dbs.Execute "ALTER TABLE receipts ADD COLUMN [Дата] DATE"

    dbs.Execute "UPDATE receipts SET [Дата] = [txtDDate];"
    
    Set dbs = Nothing
    
    accApp.Quit
End Function

