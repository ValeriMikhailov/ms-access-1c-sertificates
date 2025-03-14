Attribute VB_Name = "mdlLinksToScanUTSI"
Option Compare Database
Option Explicit

Function LinksToScanUTSI()
    pathToBase = pathMainBase

    If Len(Dir$(pathToBase & fileNameLinksToUTSI)) > 0 Then
        Kill pathToBase & fileNameLinksToUTSI
        Call CreateNewDatabase(pathToBase & fileNameLinksToUTSI)
    Else
        Call CreateNewDatabase(pathToBase & fileNameLinksToUTSI)
    End If
    
    Call CreateTableLinksToScansRegUd(pathToBase & fileNameLinksToUTSI)  'создание таблицы под резутаты парсинга
    Call colKeysLinksUTSI                                       'парсинг

    DoCmd.OpenQuery qryPathUtsi, , acEdit
End Function
