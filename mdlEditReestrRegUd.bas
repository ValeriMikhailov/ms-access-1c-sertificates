Attribute VB_Name = "mdlEditReestrRegUd"
Option Compare Database
Option Explicit

Function editRegUdBase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
'удаление добавленной строки
    Call crtRegUdSQLdelRow
    
    Dim dbs As Database
    Set dbs = OpenDatabase(pathRegUdBase & fileNameRegUdBase)
'разборка полей содержащих ƒј“џ
'    dbs.Execute "ALTER TABLE dbRegUd ADD COLUMN [registration_date] DATE"
'    dbs.Execute "ALTER TABLE dbRegUd ADD COLUMN [registration_date_end] DATE"

'    dbs.Execute "UPDATE dbRegUd SET [registration_date] = RIGHT([txtDBegin],10);"
'    dbs.Execute "UPDATE dbRegUd SET [registration_date_end] = RIGHT([txtDEnd],10);"
'    dbs.Execute "UPDATE dbRegUd SET [registration_date_end] = #1/1/2021# WHERE RIGHT([txtDEnd],4) = '1416';"
    
'    dbs.Execute "UPDATE dbRegUd SET [okp] = REPLACE([okp],Chr(160),'');" 'неразрывный пробел
'    dbs.Execute "UPDATE dbRegUd SET [okp] = REPLACE([okp],Chr(32),'');" 'пробел
'    dbs.Execute "UPDATE dbRegUd SET [okp] = REPLACE([okp],Chr(9),'');" 'знак табул€ции
    
'    dbs.Execute "UPDATE dbRegUd SET [name] = REPLACE([name],'<br>', Chr(32));"
'    dbs.Execute "UPDATE dbRegUd SET [name] = REPLACE([name],Chr(160), Chr(32));"
'    dbs.Execute "UPDATE dbRegUd SET [name] = REPLACE([name],Chr(9), Chr(32));"
'    dbs.Execute "UPDATE dbRegUd SET [name] = REPLACE([name],Chr(32)&Chr(32), Chr(32));"
    
'    dbs.Execute "UPDATE dbRegUd SET [producer] = REPLACE([producer],Chr(34), '');" 'кавычки
    
    Set dbs = Nothing
    
    accApp.Quit
End Function
