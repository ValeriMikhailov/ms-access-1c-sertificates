Attribute VB_Name = "mdlLinksToScansResult"
Option Compare Database
Option Explicit

Function LinksUpdateToScansResult()
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameLinksToScansRegUd)
'��������� �������� ���� � �������������� �������
    dbs.Execute "ALTER TABLE RUsearchResult ALTER COLUMN pathNum TEXT PRIMARY KEY"
'��������� � �������������� ������� ������ �� ������� �� ������� �����
    dbs.Execute "INSERT INTO RUsearchResult(pathNum) SELECT [" & tblNameRegUd1C & "].pathNum " & _
                "FROM [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "]"
'��������� � �������������� ������� ������ ���� � inf.pdf
    dbs.Execute "UPDATE RUsearchResult " & _
                "SET [" & tblNameRegUdResult & "].FPath = Chr(35) & '" & pathToServer & fileInfo & "' & Chr(35) " & _
                "WHERE (([" & tblNameRegUdResult & "].FPath) Is Null)"
                
    dbs.Close
End Function
