Attribute VB_Name = "mdlColKeysLinksRegUd"
Option Compare Database
Option Explicit

Function colKeysLinksRegUd()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "�� *.*"
'    colSearchKeys.Add "* �� *.*"
    colSearchKeys.Add "*+�� *.*"
    
    pathRu colSearchKeys
End Function
