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
    
    CreateTableLinksToScansRegUd pathToBase & fileNameLinksToScansRegUd '�������� ������� ��� ���������� ��������
    
    Call colKeysLinksRegUd                      '�������
    Call crtListKeysTableName                   '�������� ������� listKeysRegUd
    
'    exportToExcel pathMainBase & fileNameLinksToScansRegUd             '������� ������� � tmpExpAcc.xls
'    DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel8, tableName:=listKeysTableName, FileName:=pathMainBase & exportFileName   '������� ������� � tmpExpAcc.xls
'    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   '������ Excel linkActions

    pathToBase = pathMainBase
    fileNameBase = fileNameLinksToScansRegUd
    fileExcelImport = fileNameExcelTemplate
    tableNameAction = linksTableNameGet
    
'    nameExcelMacro = nameExcelMacroLinksDeleteSheets                                '�������� ������ � Excel
'    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)    '������ ������� RUsearch
'    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   '������ Excel linkActions
'    Kill pathToBase & exportFileName                                  '�������� tmpExpAcc.xls
    
'����������� ������ � ������� Access "RUsearchResult"
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameLinksToScansRegUd)
'    dbs.Execute "INSERT INTO RUsearchResult(pathNum, FPath) SELECT RUsearch.pathNum, RUsearch.FPath FROM RUsearch"
    dbs.Execute "INSERT INTO RUsearchResult(pathNum, FPath) SELECT dataRegUd.pathNum, RU.FPath FROM RU, [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] WHERE (((InStr(1,[FPath],[pathNum] & Chr(32)))>0));"
    dbs.Close
    
'TODO ���������� ������� �� �������� "RUsearchResult"
    Call crtLinksDoublesRUsearch
    DoCmd.OpenQuery qryLinksDoublesRUsearch, , acEdit    '�������� ������� ������, "RUsearchResult"
    
    Call crtLinksRUsearch
    DoCmd.OpenQuery qryLinksRUsearch, , acEdit
End Function
