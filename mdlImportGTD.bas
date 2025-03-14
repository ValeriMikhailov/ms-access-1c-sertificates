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
    GTDbaseSQL = "SELECT [" & tabNameGTDbase & "].[Код], [" & tabNameGTDbase & "].[ГТД], " & _
            "[Наименование], [Производитель], [Основной поставщик], [поставщикКод], [turnover] AS [обр], [Количество] AS [к-во], [ЕдХран] AS [хран] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameGTDbase & "].[" & tabNameGTDbase & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameGTDbase & "].Код " & _
            "WHERE ((([" & tabNameMainBase & "].[Код группы]) Not Like '00000074367' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00010002852' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00000074501' And ([" & tabNameMainBase & "].[Код группы]) Not Like '000074503' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00010001066' And ([" & tabNameMainBase & "].[Код группы]) Not Like '000074636' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00010000936' And ([" & tabNameMainBase & "].[Код группы]) Not Like '000074500' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00010004072' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00000028543' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00000000001' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00000023949' And ([" & tabNameMainBase & "].[Код группы]) Not Like '00010016371')) " & _
            "ORDER BY Наименование;"
    Call crtQuery(GTDbaseSQL, qryOpenGTDbase)
    DoCmd.OpenQuery qryOpenGTDbase, acViewNormal
End Function
Function crtGTDbaseSQLDelRow()
    GTDbaseSQLDelRow = "DELETE " & _
                        "[" & tabNameGTDbase & "].[Код] " & _
                    "FROM [" & pathMainBase & fileNameGTDbase & "].[" & tabNameGTDbase & "] " & _
                    "WHERE [Код]='a'"
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
    
    dbs.Execute "UPDATE [GTD] SET [ГТД] = 0 WHERE [ГТД] = 'Нет';"
    dbs.Execute "UPDATE [GTD] SET [ГТД] = 1 WHERE [ГТД] = 'Да';"
    
    Set dbs = Nothing
    accApp.Quit
End Function

