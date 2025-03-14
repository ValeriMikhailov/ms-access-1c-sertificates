Attribute VB_Name = "mdlColKeysLinksUTSI"
Option Compare Database
Option Explicit

Function colKeysLinksUTSI()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "срях *.*"
    
    pathRu colSearchKeys
End Function
