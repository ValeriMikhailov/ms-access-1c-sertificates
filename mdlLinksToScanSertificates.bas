Attribute VB_Name = "mdlLinksToScanSertificates"
Option Compare Database
Option Explicit

Function LinksToScanSertificates()
    pathToBase = pathMainBase

    If Len(Dir$(pathToBase & fileNameLinksToScansSertificate)) > 0 Then
        Kill pathToBase & fileNameLinksToScansSertificate
        Call CreateNewDatabase(pathToBase & fileNameLinksToScansSertificate)
    Else
        Call CreateNewDatabase(pathToBase & fileNameLinksToScansSertificate)
    End If
    
    Call CreateTableLinksToScansRegUd(pathToBase & fileNameLinksToScansSertificate)  '�������� ������� ��� �������� ��������
    Call colKeysLinksSertificates                                       '�������

    crtPathSertificates
End Function
