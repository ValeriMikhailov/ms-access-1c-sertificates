Attribute VB_Name = "mdlEditMainBase"
Option Compare Database
Option Explicit

Function editMainBase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    crtMainBaseSQLDelRow
    
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameMainBase)
'�������� ����� ���������� ����
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [��������] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [���� ��] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [���� �������� ��] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [���� �������� ���������� ������������] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [���� �������� ���������� ������������] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [���� ��] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [�����] DATE"
    
    dbs.Execute "UPDATE mainNomenclature SET [��������] = LEFT([txtDCard],10);"
    dbs.Execute "UPDATE mainNomenclature SET [���� ��] = LEFT([txtDBgnRegU],10);"
    dbs.Execute "UPDATE mainNomenclature SET [���� �������� ��] = LEFT([txtDEndRegU],10);"
    dbs.Execute "UPDATE mainNomenclature SET [���� �������� ���������� ������������] = LEFT([txtDSs],10);"
    dbs.Execute "UPDATE mainNomenclature SET [���� �������� ���������� ������������] = LEFT([txtDDs],10);"
    dbs.Execute "UPDATE mainNomenclature SET [���� ��] = LEFT([txtUTSI],10);"
    dbs.Execute "UPDATE mainNomenclature SET [�����] = LEFT([txtDBill],10);"
' ��/���
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [������� ��������] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [������� ��������] = 1 WHERE [txtDel] = '���';"
    dbs.Execute "UPDATE mainNomenclature SET [������� ��������] = 2 WHERE [txtDel] = '��';"
    
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [turnover] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [turnover] = 0 WHERE [txtTurnover] = '���';"
    dbs.Execute "UPDATE mainNomenclature SET [turnover] = 1 WHERE [txtTurnover] = '��';"
    
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [trace] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [trace] = 0 WHERE [txtTrace] = '���';"
    dbs.Execute "UPDATE mainNomenclature SET [trace] = 1 WHERE [txtTrace] = '��';"

'    dbs.Execute "ALTER TABLE mainNomenclature DROP COLUMN textDate DATE"
'    dbs.Execute "ALTER TABLE mainNomenclature ALTER COLUMN [���] TEXT(11)"
    
    Set dbs = Nothing
'���������� ���������� ������� ��� ���������� � ��� ����
    crtDoubleRows
    DoCmd.Close
    crtDoubleRowsDel
    DoCmd.Close

    fncDelDoubleRow '�������� ������ ��������
    crtMainBaseSQLtoDescr '���������� ����� ����� �������� � Descr.accdb
    crtMainBaseSQLtoRegUd1C '���������� "����� ��" � dbRegUd1C.accdb
    accApp.Quit
End Function
Sub editNames()
    Dim dbs As Database
    Set dbs = CurrentDb()

    dbs.Execute "UPDATE main SET [������������] = REPLACE([������������],Chr(10),'');" 'Linefeed character ������ �������� ������
    dbs.Execute "UPDATE main SET [�����] = REPLACE([�����],Chr(10),'');" 'Linefeed character ������ �������� ������
'    dbs.Execute "UPDATE main SET [���] = REPLACE([���],Chr(10),'');" '��� ���� �� �������� � �������
End Sub

