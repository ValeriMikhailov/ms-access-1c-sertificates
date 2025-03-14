Attribute VB_Name = "mdlLinksToScansRegUd"
Option Compare Database
Option Explicit

Function LinksToScansRegUd()
'    Dim excelApp
'    nameExcelMacro = nameExcelMacroLinksRegUd
    pathToBase = pathMainBase

    If Len(Dir$(pathToBase & fileNameLinksToScansRegUd)) > 0 Then
        Kill pathToBase & fileNameLinksToScansRegUd
        CreateNewDatabase pathToBase & fileNameLinksToScansRegUd
    Else
        CreateNewDatabase pathToBase & fileNameLinksToScansRegUd
    End If
    
    CreateTableLinksToScansRegUd pathToBase & fileNameLinksToScansRegUd 'создание таблицы под результаты парсинга
    
    Call colKeysLinksRegUd                      'парсинг
    Call crtListKeysTableName                   'создание запроса listKeysRegUd
    
'    exportToExcel pathMainBase & fileNameLinksToScansRegUd             'экспорт таблицы в tmpExpAcc.xls
'    DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel8, tableName:=listKeysTableName, FileName:=pathMainBase & exportFileName   'экспорт запроса в tmpExpAcc.xls
'    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   'макрос Excel linkActions

    pathToBase = pathMainBase
    fileNameBase = fileNameLinksToScansRegUd
    fileExcelImport = fileNameExcelTemplate
    tableNameAction = linksTableNameGet
    
'    nameExcelMacro = nameExcelMacroLinksDeleteSheets                                'удаление листов в Excel
'    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)    'импорт таблицы RUsearch
'    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   'макрос Excel linkActions
'    Kill pathToBase & exportFileName                                  'удаление tmpExpAcc.xls
    
'копирование данных в таблицу Access "RUsearchResult"
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameLinksToScansRegUd)
'    dbs.Execute "INSERT INTO RUsearchResult(pathNum, FPath) SELECT RUsearch.pathNum, RUsearch.FPath FROM RUsearch"
    dbs.Execute "INSERT INTO RUsearchResult(pathNum, FPath) SELECT dataRegUd.pathNum, RU.FPath FROM RU, [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] WHERE (((InStr(1,[FPath],[pathNum] & Chr(32)))>0));"
    dbs.Close
    
'TODO переписать запросџ на открытие "RUsearchResult"
    Call crtLinksDoublesRUsearch
    DoCmd.OpenQuery qryLinksDoublesRUsearch, , acEdit    'открытие запроса ƒ”ЅЋ≈…, "RUsearchResult"
    
    Call crtLinksRUsearch
    DoCmd.OpenQuery qryLinksRUsearch, , acEdit
End Function
