Attribute VB_Name = "mdlColKeysLinksUTSI"
Option Compare Database
Option Explicit

Function colKeysLinksUTSI()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "���� *.*"
    
    pathRu colSearchKeys
End Function
