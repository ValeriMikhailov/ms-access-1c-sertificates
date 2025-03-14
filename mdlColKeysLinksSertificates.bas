Attribute VB_Name = "mdlColKeysLinksSertificates"
Option Compare Database
Option Explicit

Function colKeysLinksSertificates()
    Dim colSearchKeys As New Collection
    colSearchKeys.Add "ÄÑ *.*"
    colSearchKeys.Add "ÑÑ *.*"
    colSearchKeys.Add "* ÄÑ *.*"
    colSearchKeys.Add "* ÑÑ *.*"
    colSearchKeys.Add "*+ÄÑ *.*"
    colSearchKeys.Add "*+ÑÑ *.*"
    colSearchKeys.Add "ÓÒÑÈ *.*"
    colSearchKeys.Add "îòêàç *.*"
    
    pathRu colSearchKeys
End Function
