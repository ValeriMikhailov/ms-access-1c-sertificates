Attribute VB_Name = "mdlImportReceipts"
Option Compare Database
Option Explicit

Function StartCreateNewReceipts()
'    Dim excelApp
    pathToBase = pathMainBase
    fileNameBase = fileNameReceipts
    fileExcelImport = fileExcelReceipts
    tableNameAction = tabReceipts
'    nameExcelMacro = nameExcelMacroReceipts
    
    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelReceipts
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelReceipts
        CreateNewDatabase pathToBase & fileNameBase
    End If
    
    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)
    editReceipts pathMainBase & fileNameReceipts
    Call crtReceiptsQry
End Function

