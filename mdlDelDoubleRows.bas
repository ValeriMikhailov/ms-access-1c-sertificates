Attribute VB_Name = "mdlDelDoubleRows"
Option Compare Database
Option Explicit

Function fncDelDblRows()
    Dim dbs As Database
    Dim rstDynasetDoubleRows, rstLoopDoubleRows As Recordset
    Dim strID As String

    Set dbs = CurrentDb()
    With dbs
        Set rstDynasetDoubleRows = .OpenRecordset(qryDoubleRows, dbOpenDynaset)
    '        If rstDynasetDoubleRows.EOF Then Exit Function
            With rstDynasetDoubleRows
                Do Until .EOF
                    strID = .[Код]
                    fncDelDblRows = strID
                    .MoveNext
                Loop
            End With
            rstDynasetDoubleRows.Close
        .Close
    End With
End Function
Function fncDelDoubleRow()
    Dim dbs As Database
    Dim rstDynasetDoubleRows As Recordset
    Dim strID As String

    Set dbs = CurrentDb()
    With dbs
        Set rstDynasetDoubleRows = .OpenRecordset(qryDoubleRows, dbOpenDynaset)
    '        If rstDynasetDoubleRows.EOF Then Exit Function
            With rstDynasetDoubleRows
                Do Until .EOF
                    dbs.Execute "DELETE * FROM [" & qrydoubleRowsDel & "] WHERE (([" & qrydoubleRowsDel & "].[Код])=fncDelDblRows());"
                    .MoveNext
                Loop
            End With
            rstDynasetDoubleRows.Close
        .Close
    End With
    altCodID
End Function
Function altCodID()
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameMainBase)
    dbs.Execute "ALTER TABLE [" & tabNameMainBase & "] ALTER COLUMN [Код] TEXT PRIMARY KEY;"
    dbs.Close
End Function
