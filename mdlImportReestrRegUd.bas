Attribute VB_Name = "mdlImportReestrRegUd"
Option Compare Database
Option Explicit

Function ImportRegUd()
    pathToBase = pathRegUdBase
    fileNameBase = fileNameRegUdBase
    fileExcelImport = fileExcelReestRegUd
    tableNameAction = tabReestRegUd

    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelRegUd
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelRegUd
        CreateNewDatabase pathToBase & fileNameBase
    End If

    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)
    
    editRegUdBase pathToBase & fileNameBase
    crtRegUdSearch
End Function
