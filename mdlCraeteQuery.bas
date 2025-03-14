Attribute VB_Name = "mdlCraeteQuery"
Option Compare Database
Option Explicit

Function crtQuery(strSQL As String, strQryName As String)
    Dim dbs As Database
    Dim qdf As QueryDef
    Set dbs = CurrentDb
' опрделяем существует запрос или нет, если да то удаляем его
    dbs.QueryDefs.Refresh
    DoCmd.SetWarnings False
    For Each qdf In dbs.QueryDefs
        If qdf.Name = strQryName Then
            dbs.QueryDefs.Delete qdf.Name
        End If
    Next qdf
' создаем QueryDef
    Set qdf = dbs.CreateQueryDef(strQryName, strSQL)
' открываем
'    DoCmd.OpenQuery strQryName, acViewNormal
    qdf.Close
    Set qdf = Nothing
End Function
