Attribute VB_Name = "mdlColKeysLinksRegUd"
Option Compare Database
Option Explicit

Function colKeysLinksRegUd()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "аг *.*"
'    colSearchKeys.Add "* аг *.*"
    colSearchKeys.Add "*+аг *.*"
    
    pathRu colSearchKeys
End Function
