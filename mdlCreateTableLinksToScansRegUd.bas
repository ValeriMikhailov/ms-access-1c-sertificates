Attribute VB_Name = "mdlCreateTableLinksToScansRegUd"
Option Compare Database
Option Explicit

'TODO переписать на переменные
Function CreateTableLinksToScansRegUd(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
'Создание Таблицы "RU"--------------------------------------------------------------------
'    Const strTableName As String = "RU" 'Название таблицы Access
    Dim tbl As TableDef       'объект таблица
    Dim fld As Field          'объект поле
'    Dim rst As Recordset      'объект набор записей
'создание объектной переменной таблицы, полей /// и индекса в ней -----------------------
    Set tbl = accApp.CurrentDb.CreateTableDef(strTableName)
    Set fld = tbl.CreateField(strFieldName, dbMemo)
    fld.Attributes = dbHyperlinkField
    
    With tbl.Fields
        .Append fld
    End With
'Фактическое добавление таблицы из объектной переменной описанной выше --------------
    accApp.CurrentDb.TableDefs.Append tbl
    
'Создание Таблицы "RUsearchResult"--------------------------------------------------------------------
    Const strTableNameResult As String = tblNameRegUdResult 'Название таблицы Access
    Set tbl = accApp.CurrentDb.CreateTableDef(strTableNameResult)
    Set fld = tbl.CreateField("pathNum", dbText)
    
    With tbl.Fields
        .Append fld
    End With
    
    Set fld = tbl.CreateField(strFieldName, dbMemo)
    fld.Attributes = dbHyperlinkField
    
    With tbl.Fields
        .Append fld
    End With
'Фактическое добавление таблицы из объектной переменной описанной выше --------------
    accApp.CurrentDb.TableDefs.Append tbl
    
'    Set accApp = Nothing
    accApp.Quit
End Function
