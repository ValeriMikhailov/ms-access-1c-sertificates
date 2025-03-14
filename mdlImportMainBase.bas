Attribute VB_Name = "mdlImportMainBase"
Option Compare Database
Option Explicit

Function StartCreateNewDBandImportExcelTable()
'    Dim excelApp
    pathToBase = pathMainBase
    fileNameBase = fileNameMainBase
    fileExcelImport = fileExcelMainBase
    tableNameAction = tabNameMainBase
'    nameExcelMacro = nameExcelMacroImpMainBase
    
    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelMainBase
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelMainBase
        CreateNewDatabase pathToBase & fileNameBase
    End If
    
    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)

    editMainBase pathToBase & fileNameBase
    Call crtUPDATE
    DoCmd.OpenQuery qryOpenMainNomenclature, , acEdit
'    DoCmd.OpenQuery "doubleRows", , acEdit
End Function
