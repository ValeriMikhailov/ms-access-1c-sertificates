Attribute VB_Name = "mdlImportGTD"
Option Compare Database
Option Explicit

Public Const fileNameGTDbase As String = "GTD.accdb"
Public Const tabNameGTDbase As String = "GTD"
Public Const qryOpenGTDbase As String = "GTDbase"
Public GTDbaseSQL As String
Public GTDbaseSQLDelRow As String
Public Const qryGTDbaseDelRow As String = "GTDbaseDelRow"

Function StartCreateNewGTDbase()
    pathToBase = pathMainBase
    fileNameBase = fileNameGTDbase
    fileExcelImport = fileNameExcelGTD
    tableNameAction = tabNameGTDbase
    
    If Len(Dir$(pathToBase & fileNameBase)) > 0 Then
        Kill pathToBase & fileNameBase
        Call EditExcelGTD
        CreateNewDatabase pathToBase & fileNameBase
    Else
        Call EditExcelGTD
        CreateNewDatabase pathToBase & fileNameBase
    End If
    
    ImportExcelTable pathToBase & fileNameBase, impExcelTab(pathToBase & fileNameBase, tableNameAction, fileExcelImport)
    editGTDbase pathToBase & fileNameBase
    DoCmd.OpenQuery qryOpenGTDbase, , acEdit
End Function
Function crtGTDbase()
    GTDbaseSQL = "SELECT [" & tabNameGTDbase & "].[���], [" & tabNameGTDbase & "].[���], " & _
            "[������������], [�������������], [�������� ���������], [������������], [turnover] AS [���], [����������] AS [�-��], [������] AS [����] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameGTDbase & "].[" & tabNameGTDbase & "] " & _
                    "ON [" & tabNameMainBase & "].��� = [" & tabNameGTDbase & "].��� " & _
            "WHERE ((([" & tabNameMainBase & "].[��� ������]) Not Like '00000074367' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00010002852' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00000074501' And ([" & tabNameMainBase & "].[��� ������]) Not Like '000074503' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00010001066' And ([" & tabNameMainBase & "].[��� ������]) Not Like '000074636' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00010000936' And ([" & tabNameMainBase & "].[��� ������]) Not Like '000074500' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00010004072' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00000028543' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00000000001' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00000023949' And ([" & tabNameMainBase & "].[��� ������]) Not Like '00010016371')) " & _
            "ORDER BY ������������;"
    Call crtQuery(GTDbaseSQL, qryOpenGTDbase)
    DoCmd.OpenQuery qryOpenGTDbase, acViewNormal
End Function
Function crtGTDbaseSQLDelRow()
    GTDbaseSQLDelRow = "DELETE " & _
                        "[" & tabNameGTDbase & "].[���] " & _
                    "FROM [" & pathMainBase & fileNameGTDbase & "].[" & tabNameGTDbase & "] " & _
                    "WHERE [���]='a'"
    Call crtQuery(GTDbaseSQLDelRow, qryGTDbaseDelRow)
    DoCmd.OpenQuery qryGTDbaseDelRow, acViewNormal
End Function
Function editGTDbase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    crtGTDbaseSQLDelRow
        
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameGTDbase)
    
    dbs.Execute "UPDATE [GTD] SET [���] = 0 WHERE [���] = '���';"
    dbs.Execute "UPDATE [GTD] SET [���] = 1 WHERE [���] = '��';"
    
    Set dbs = Nothing
    accApp.Quit
End Function

