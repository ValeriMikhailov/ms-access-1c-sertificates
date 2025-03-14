Attribute VB_Name = "mdlSearchOK"
Option Compare Database
Option Explicit

Function SearchOK() As Boolean
    sSearch = InputBox("Please, insert 'z' for searching links RegUd or nothing for searchin links to scans Sertificates")
    If sSearch = Format(Now, sParameterSearch) Then
        SearchOK = True
    End If
End Function
