Attribute VB_Name = "mdlImportSize"
Option Compare Database
Option Explicit

'Public Const fileNameSize As String = "size.accdb"
Public Const tabNameSize As String = "size"
Public sizeSQLDelRow As String
Public Const qrySizeSQLDelRow As String = "SizeSQLDelRow"

Function StartCreateNewSize()
    pathToBase = pathMainBase
    fileNameBase = fileNameSize
    fileExcelImport = "size.xlsx"
    tableNameAction = "size"
    
    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelSize
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelSize
        CreateNewDatabase pathToBase & fileNameBase
    End If
    
    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)
    editSizeBase pathToBase & fileNameBase
    crtCountPlaceis
    DoCmd.Close
    DoCmd.OpenForm "sizeBrutto", acNormal
End Function
Function crtSizeSQLDelRow()
    sizeSQLDelRow = "DELETE " & _
                        "[" & tabNameSize & "].[Код] " & _
                    "FROM [" & pathMainBase & fileNameSize & "].[" & tabNameSize & "] " & _
                    "WHERE [Код]='a'"
    Call crtQuery(sizeSQLDelRow, qrySizeSQLDelRow)
    DoCmd.OpenQuery qrySizeSQLDelRow, acViewNormal
End Function
Function editSizeBase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    crtSizeSQLDelRow
'    Set accApp = Nothing
    accApp.Quit
End Function
