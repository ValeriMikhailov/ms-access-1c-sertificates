Attribute VB_Name = "mdlColKeysLinksSertificates"
Option Compare Database
Option Explicit

Function colKeysLinksSertificates()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "�� *.*"
    colSearchKeys.Add "�� *.*"
    colSearchKeys.Add "* �� *.*"
    colSearchKeys.Add "* �� *.*"
    colSearchKeys.Add "*+�� *.*"
    colSearchKeys.Add "*+�� *.*"
    colSearchKeys.Add "���� *.*"
    colSearchKeys.Add "����� *.*"
    
    pathRu colSearchKeys
End Function
