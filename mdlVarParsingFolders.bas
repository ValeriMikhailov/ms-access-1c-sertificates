Attribute VB_Name = "mdlVarParsingFolders"
Option Compare Database
Option Explicit

'    Public colSearchKeys As New Collection
' ------------ making links to scans RegUd ----------------
    Public Const fileNameLinksToScansRegUd As String = "LinksToScansRegUd.accdb"   '�������� ����� � ������������ ��������
    Public Const tblNameRegUdResult As String = "RUsearchResult"
    
    Public Const fileNameLinksToScansSertificate As String = "LinksToScansSertificates.accdb"   '�������� ����� � ������������ ��������
'    Public Const fileNameLinksToUTSI As String = "LinksToScansUTSI.accdb"
    
    Public Const exportFileName As String = "tmpExpAcc.xls"        '�������� ����� � ������� ������������� ���������� ��������
    Public Const fileNameExcelTemplate As String = "Template_RegUd.xlsx"  '�������� �����
    Public Const linksTableName As String = "RU"                   '����� �������� ������� ������ ��������� � ��������� ������� � ������� "linkActions"
    Public Const linksTableNameGet As String = "RUsearch"                   '����� �������� ������� ������ ��������� � ��������� ������� � ������� "linkActions"
    Public Const listKeysTableName As String = "listKeysRegUd"     '����� �������� ������� ������ ��������� � ��������� ������� � ������� "linkActions"
    
    Public Const nameExcelMacroLinksRegUd As String = "linkActions"  '�������� ������� Excel ������������� ����
    Public Const nameExcelMacroLinksDeleteSheets As String = "linkActionsSheetsDelete"  '�������� ������� Excel ������������� ����
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

