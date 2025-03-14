Attribute VB_Name = "mdlVarParsingFolders"
Option Compare Database
Option Explicit

'    Public colSearchKeys As New Collection
' ------------ making links to scans RegUd ----------------
    Public Const fileNameLinksToScansRegUd As String = "LinksToScansRegUd.accdb"   'Название файла с результатами парсинга
    Public Const tblNameRegUdResult As String = "RUsearchResult"
    
    Public Const fileNameLinksToScansSertificate As String = "LinksToScansSertificates.accdb"   'Название файла с результатами парсинга
'    Public Const fileNameLinksToUTSI As String = "LinksToScansUTSI.accdb"
    
    Public Const exportFileName As String = "tmpExpAcc.xls"        'Название файла в который экпортируются результаты парсинга
    Public Const fileNameExcelTemplate As String = "Template_RegUd.xlsx"  'Название файла
    Public Const linksTableName As String = "RU"                   'ВАЖНО Название таблицы должно совпадать с названием таблицы в макросе "linkActions"
    Public Const linksTableNameGet As String = "RUsearch"                   'ВАЖНО Название таблицы должно совпадать с названием таблицы в макросе "linkActions"
    Public Const listKeysTableName As String = "listKeysRegUd"     'ВАЖНО Название таблицы должно совпадать с названием таблицы в макросе "linkActions"
    
    Public Const nameExcelMacroLinksRegUd As String = "linkActions"  'Название макроса Excel перемещающего лист
    Public Const nameExcelMacroLinksDeleteSheets As String = "linkActionsSheetsDelete"  'Название макроса Excel перемещающего лист
    Public Const nameExcelMacroFormatingCells As String = "changesForEnd"
    
'========///======== PARSING =========///============
    Public strPath As String
    Public booIncludeSubfolders As Boolean
    Public strFileSpec As Variant
    Public strFileSpecLoop As String
    Public strTemp As String
    Public vFolderName As Variant
    Public strSQL As String
    Public strFolder As String
    Public strNameTblForPathsScans As String
    Public sSearch As String

    Public gCount As Long ' added by Crystal
    Public Const sParameterSearch As String = "z"
    Public Const strTableName As String = "RU"
    Public Const strFieldName As String = "FPath"

'========///======== Forms =========///============
    Public Const frmnameOtkaz = "Otkaz"
    Public Const frmnamePathSertificates = "PathSertificates"
    Public Const frmnamePathRegUd = "PathRegUd"

