Attribute VB_Name = "mdlCreateTableLinksToScansRegUd"
Option Compare Database
Option Explicit

'TODO ���������� �� ����������
Function CreateTableLinksToScansRegUd(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
'�������� ������� "RU"--------------------------------------------------------------------
'    Const strTableName As String = "RU" '�������� ������� Access
    Dim tbl As TableDef       '������ �������
    Dim fld As Field          '������ ����
'    Dim rst As Recordset      '������ ����� �������
'�������� ��������� ���������� �������, ����� /// � ������� � ��� -----------------------
    Set tbl = accApp.CurrentDb.CreateTableDef(strTableName)
    Set fld = tbl.CreateField(strFieldName, dbMemo)
    fld.Attributes = dbHyperlinkField
    
    With tbl.Fields
        .Append fld
    End With
'����������� ���������� ������� �� ��������� ���������� ��������� ���� --------------
    accApp.CurrentDb.TableDefs.Append tbl
    
'�������� ������� "RUsearchResult"--------------------------------------------------------------------
    Const strTableNameResult As String = tblNameRegUdResult '�������� ������� Access
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
'����������� ���������� ������� �� ��������� ���������� ��������� ���� --------------
    accApp.CurrentDb.TableDefs.Append tbl
    
'    Set accApp = Nothing
    accApp.Quit
End Function
