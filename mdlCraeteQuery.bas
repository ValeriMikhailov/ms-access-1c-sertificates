Attribute VB_Name = "mdlCraeteQuery"
Option Compare Database
Option Explicit

Function crtQuery(strSQL As String, strQryName As String)
    Dim dbs As Database
    Dim qdf As QueryDef
    Set dbs = CurrentDb
' ��������� ���������� ������ ��� ���, ���� �� �� ������� ���
    dbs.QueryDefs.Refresh
    DoCmd.SetWarnings False
    For Each qdf In dbs.QueryDefs
        If qdf.Name = strQryName Then
            dbs.QueryDefs.Delete qdf.Name
        End If
    Next qdf
' ������� QueryDef
    Set qdf = dbs.CreateQueryDef(strQryName, strSQL)
' ���������
'    DoCmd.OpenQuery strQryName, acViewNormal
    qdf.Close
    Set qdf = Nothing
End Function
