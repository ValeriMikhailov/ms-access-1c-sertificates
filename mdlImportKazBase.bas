Attribute VB_Name = "mdlImportKazBase"
Option Compare Database
Option Explicit

Public Const fileNameKazBase As String = "kazDB.accdb"
Public Const tabNameKazBase As String = "kazBase"
Public Const qryOpenKazBase As String = "kazakhstan"
Public kazBaseSQL As String
Public kazBaseSQLDelRow As String
Public Const qryKazBaseDelRow As String = "KazBaseDelRow"
Public Const qryAuthorKaz = "authorKaz"
Public authorKazSQL As String

Function StartCreateNewKazBase()
    pathToBase = pathMainBase
    fileNameBase = fileNameKazBase
    fileExcelImport = fileNameExcelKazBase
    tableNameAction = tabNameKazBase
    
    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelKazBase
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelKazBase
        CreateNewDatabase pathToBase & fileNameBase
    End If
    
    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)
    editKazBase pathToBase & fileNameBase
    DoCmd.OpenQuery qryOpenKazBase, , acEdit
End Function
Function crtOpenKazBase()
    kazBaseSQL = "SELECT cod, kazBase.groupID, kazBase.grName, wName, pName, NDS, unit, unit_st, dateCard, descrip " & _
        "FROM [" & pathMainBase & fileNameKazBase & "].[" & tabNameKazBase & "]" & _
        "WHERE ([pName] IS NOT NULL);"
    Call crtQuery(kazBaseSQL, qryOpenKazBase)
    DoCmd.OpenQuery qryOpenKazBase, acViewNormal
End Function
Function crtKazBaseSQLDelRow()
    kazBaseSQLDelRow = "DELETE " & _
                        "[" & tabNameKazBase & "].[cod] " & _
                    "FROM [" & pathMainBase & fileNameKazBase & "].[" & tabNameKazBase & "] " & _
                    "WHERE [cod]='a'"
    Call crtQuery(kazBaseSQLDelRow, qryKazBaseDelRow)
    DoCmd.OpenQuery qryKazBaseDelRow, acViewNormal
End Function
Function editKazBase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    crtKazBaseSQLDelRow
        
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameKazBase)
    
    dbs.Execute "ALTER TABLE [kazBase] ADD COLUMN dateCard DATE"
    dbs.Execute "UPDATE [kazBase] SET dateCard = textDate;"
'    dbs.Execute "UPDATE [kazBase] SET dateCard = #1/1/1911# WHERE [kazBase].[dateCard] Is Null;"
    dbs.Execute "ALTER TABLE [kazBase] DROP COLUMN textDate DATE"
    dbs.Execute "ALTER TABLE [kazBase] ALTER COLUMN wName TEXT(150)"
    
    Set dbs = Nothing
    accApp.Quit
End Function
Function crtAuthorKazQry()
    authorKazSQL = "SELECT " & _
                    "cod, " & _
                    "dateCard " & _
                "FROM [" & pathMainBase & fileNameKazBase & "].[" & tabNameKazBase & " ] " & _
                "WHERE ((([" & tabNameKazBase & " ].dateCard) " & _
                    "Between " & SD & " " & _
                    "And " & FD & ") " & _
                    "AND (([" & tabNameKazBase & "].author)='Михайлов Валерий')) " & _
                "ORDER BY [" & tabNameKazBase & " ].dateCard DESC"
    Call crtQuery(authorKazSQL, qryAuthorKaz)
    DoCmd.OpenQuery qryAuthorKaz, , acEdit
End Function

