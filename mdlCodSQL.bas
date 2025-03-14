Attribute VB_Name = "mdlCodSQL"
Option Compare Database
Option Explicit
Function crtCsvRu() ' *** ВНИМАНИЕ *** Not Like '00010022145' *** по РЕШЕНИЮ РУКОВОДСТВА
    getCsvRuSQL = "SELECT " & _
                    "'/mnt' & [" & tblNameRegUdResult & "].FPath AS path, " & _
                    "'RU-' & Код & '.pdf' AS card " & _
            "FROM ([" & pathMainBase & fileNameLinksToScansRegUd & "].[" & tblNameRegUdResult & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
                    "ON [" & tblNameRegUdResult & "].pathNum = [" & tblNameRegUd1C & "].pathNum) " & _
            "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tblNameRegUd1C & "].registration_number = [" & tabNameMainBase & "].[Номер РУ] " & _
            "WHERE ((([" & tabNameMainBase & "].[Номер РУ]) Is Not Null) AND (([" & tabNameMainBase & "].[Срок действия РУ])>Date()) " & _
                "AND (([" & tabNameMainBase & "].[" & fldMainCod & "]) Not Like '00010022145')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(getCsvRuSQL, qryGetCsvRu)
'    DoCmd.OpenQuery qryGetCsvRu, acViewNormal
End Function
Function crtCsvDs()
    getCsvDsSQL = "SELECT " & _
                    "'/mnt' & endSsDs.ssdsPath AS path, " & _
                    "'DS-' & [" & tabNameMainBase & "].Код & '.pdf' AS card " & _
                "FROM endSsDs " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endSsDs.ID = [" & tabNameDescr & "].DS " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getCsvDsSQL, qryGetCsvDs)
'    DoCmd.OpenQuery qryGetCsvDs, acViewNormal
End Function
Function crtCsvSs()
    getCsvSsSQL = "SELECT " & _
                    "'/mnt' & endSsDs.ssdsPath AS path, " & _
                    "'SS-' & [" & tabNameMainBase & "].Код & '.pdf' AS card " & _
                "FROM endSsDs " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endSsDs.ID = [" & tabNameDescr & "].ID " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getCsvSsSQL, qryGetCsvSs)
'    DoCmd.OpenQuery qryGetCsvSs, acViewNormal
End Function
Function crtCsvUtsi()
    getCsvUtsiSQL = "SELECT " & _
                    "'/mnt' & endUtsi.utsiPath AS path, " & _
                    "'UT-' & [" & tabNameMainBase & "].Код & '.pdf' AS card " & _
                "FROM endUtsi " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endUtsi.ID = [" & tabNameDescr & "].UT " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getCsvUtsiSQL, qryGetCsvUtsi)
'    DoCmd.OpenQuery qryGetCsvUtsi, acViewNormal
End Function

'Function crtAuthorQry()
'    authorSQL = "SELECT " & _
'                    "Код, " & _
'                    "карточка " & _
'                "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & " ] " & _
'                "WHERE ((([" & tabNameMainBase & " ].карточка) " & _
'                    "Between " & SD & " " & _
'                    "And " & FD & ") " & _
'                    "AND (([" & tabNameMainBase & "].Автор)='Михайлов Валерий Валентинович')) " & _
'                "ORDER BY [" & tabNameMainBase & " ].карточка DESC"
'    Call crtQuery(authorSQL, qryAuthor)
'    DoCmd.OpenQuery qryAuthor, , acEdit
'End Function

Function crtArticul()
    articulSQL = "SELECT [" & tabNameMainBase & "].Код, " & _
                    "Наименование, " & _
                    "Len([Наименование]) AS s150, " & _
                    "[Наименование] & ', ' & [Артикул] AS res, " & _
                    "[примечание] AS [прим], " & _
                    "[Артикул] AS [арт] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
            "WHERE (((InStr(1,[Наименование],[Артикул]))=0) AND ((поставщикКод)='000001028')) " & _
            "ORDER BY Len([Наименование]) DESC;"
    Call crtQuery(articulSQL, qryArticul)
    DoCmd.OpenQuery qryArticul, acViewNormal
End Function
Function crtArticulSplit()
    articulSplitSQL = "SELECT [" & tabNameMainBase & "].Код, " & _
                    "Наименование, " & _
                    "Len([Наименование]) AS s150, " & _
        "SplitArticul([Наименование],0) & ', ' & [Артикул] & '_ДУБЛЬ' & SplitArticul([Наименование],1) AS res, " & _
                    "[примечание] AS [прим], " & _
                    "[Артикул] AS [арт] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
            "WHERE (((InStr(1,[Наименование],[Артикул]))=0) AND ((поставщикКод)='000001028')) " & _
            "ORDER BY Len([Наименование]) DESC;"
    Call crtQuery(articulSplitSQL, qryArticulSplit)
    DoCmd.OpenQuery qryArticulSplit, acViewNormal
End Function
Function crtSpecComplData()
    specComplDataSQL = "SELECT " & _
                    "Код, " & _
                    "complect, " & _
                    "quantity " & _
                "FROM [" & pathMainBase & fileNameSpecComplect & "].[" & tblNameSpecComplData & "] " & _
                "ORDER BY [" & tblNameSpecComplData & " ].Код DESC;"
    Call crtQuery(specComplDataSQL, qrySpecComplData)
    DoCmd.OpenQuery qrySpecComplData, , acEdit
End Function
Function crtSpecComplResult()
    specComplResultSQL = "SELECT DISTINCT " & _
                    "[" & tblNameSpecComplData & "].Код AS spec_ID, " & _
                    "mainNomenclature_1.Наименование, " & _
                    "[" & tblNameSpecComplData & "].complect & ' -' & [quantity] & [mainNomenclature.ЕдХран] & ';' AS [cadsCompl], " & _
                    "GetNextNum([complect]) & '.' & [mainNomenclature.Наименование] & ' -' & [quantity] & [mainNomenclature.ЕдХран] & ';' AS [complection]" & _
                "FROM ([" & pathMainBase & fileNameSpecComplect & "].[" & tblNameSpecComplData & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] AS mainNomenclature_1 " & _
                    "ON [" & tblNameSpecComplData & "].Код = mainNomenclature_1.Код) " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tblNameSpecComplData & "].[complect] = mainNomenclature.Код " & _
                "WHERE (((startNum())=True)) " & _
                "ORDER BY specComplData.Код DESC;"
    Call crtQuery(specComplResultSQL, qrySpecComplResult)
    DoCmd.OpenQuery qrySpecComplResult, , acReadOnly
End Function
Function crtPost()
    postSQL = "SELECT " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование, " & _
                    "[" & tblNameFirm & "].partner, " & _
                    "[" & tblNamePost & "].contact, " & _
                    "[" & tblNamePost & "].email, " & _
                    "[" & tblNamePost & "].phone " & _
                "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                "INNER JOIN ([" & pathMainBase & fileNamePost & "].[" & tblNameFirm & "] " & _
                "INNER JOIN [" & pathMainBase & fileNamePost & "].[" & tblNamePost & "] " & _
                    "ON [" & tblNameFirm & "].[поставщикКод] = [" & tblNamePost & "].[поставщикКод]) " & _
                    "ON [" & tabNameMainBase & "].[поставщикКод] = [" & tblNameFirm & "].[поставщикКод] " & _
                "WHERE ((([" & tabNameMainBase & "].turnover) = 1)) " & _
                "ORDER BY [" & tabNameMainBase & "].Код"
    Call crtQuery(postSQL, qryPost)
    DoCmd.OpenQuery qryPost, acViewNormal
End Function
Function crtReceiptsRegUd()
    receiptsRegUdSQL = "SELECT " & _
                    "[" & tabReceipts & "].Дата, " & _
                    "[" & tblNameRegUdResult & "].FPath, " & _
                    "[" & tabReceipts & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование, " & _
                    "[" & tabNameMainBase & "].[Основной поставщик], " & _
                    "[Ставка НДС] AS [НДС], " & _
                    "registration_number AS [Номер РУ], " & _
                    "[Дата РУ] " & _
            "FROM (([" & pathMainBase & fileNameLinksToScansRegUd & "].[" & tblNameRegUdResult & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
                    "ON [" & tblNameRegUdResult & "].pathNum = [" & tblNameRegUd1C & "].pathNum) " & _
            "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tblNameRegUd1C & "].registration_number = [" & tabNameMainBase & "].[Номер РУ]) " & _
            "INNER JOIN [" & pathMainBase & fileNameReceipts & "].[" & tabReceipts & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabReceipts & "].Код " & _
            "WHERE ((([" & tabNameMainBase & "].[Ставка НДС]) " & _
                    "Not Like '20%')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(receiptsRegUdSQL, qryReceiptsRegUd)
    DoCmd.OpenQuery qryReceiptsRegUd, acViewNormal
End Function
Function crtReceiptsQry()
    receiptsSQL = "SELECT " & _
                    "Дата, " & _
                    "[" & tabReceipts & "].Код, " & _
                    "[Код группы] AS кодГр, " & _
                    "Группа, " & _
                    "[" & tabNameMainBase & "].[Основной поставщик], " & _
                    "[" & tabNameMainBase & "].Наименование, " & _
                    "[Номер РУ], " & _
                    "[Ставка НДС] AS НДС, " & _
                    "ТНВЭД, " & _
                    "ОКП, " & _
                    "ОКПД2, " & _
                    "[" & tabNameDescr & "].ID, " & _
                    "[" & tabNameMainBase & "].[Срок действия сертификат соответствия] AS СС, " & _
                    "[" & tabNameDescr & "].DS, " & _
                    "[" & tabNameMainBase & "].[Срок действия декларация соответствия] AS ДС, " & _
                    "[" & tabNameDescr & "].UT " & _
                "FROM [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN ([" & pathMainBase & fileNameReceipts & "].[" & tabReceipts & "] " & _
                "RIGHT JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabReceipts & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код " & _
                "WHERE ([Пометка удаления]=1 AND (([" & tabReceipts & "].Дата) Between Date() And (Date()-33)) AND ([" & tabNameMainBase & "].[Группа]) NOT LIKE '*услуг*') " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(receiptsSQL, qryReceipts)
    DoCmd.OpenQuery qryReceipts, , acEdit
End Function
Function crtSpecComplectQry()
    specCompectSQL = "SELECT " & _
                    "Код, " & _
                    "Наименование, " & _
                    "[Текстовое описание] " & _
                "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                "WHERE ((([Текстовое описание]) Is Not Null) " & _
                    "And (([" & tabNameMainBase & "].[Код группы]) = '00010002925')) " & _
                "ORDER BY Наименование"
    Call crtQuery(specCompectSQL, qrySpecComplect)
    DoCmd.OpenQuery qrySpecComplect, , acEdit
End Function
Function crtTemplateQry()
    templateSQL = "SELECT " & _
                    "[" & tabNameMainBase & "].[" & fldMainCod & "], " & _
                    "[" & tblNameRegUd1C & "].[" & fldNameRegUd1C & "], " & _
                    "[" & tabNameDescr & "].примечание, " & _
                    "[" & tblNameRegUd1C & "].ruNumber, " & _
                    "RU.FPath " & _
                "FROM (([" & pathMainBase & "FPath.accdb].RU " & _
                "INNER JOIN [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
                    "ON RU.ID = [" & tblNameRegUd1C & "].ruNumber) " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tblNameRegUd1C & "].[" & fldNameRegUd1C & "] = [" & tabNameMainBase & "].[" & fldNameNumRegUd & "]) " & _
                "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].[" & fldMainCod & "] = [" & tabNameDescr & "].[" & fldDescrCod & "] " & _
                "WHERE (([" & tabNameMainBase & "].[Ставка НДС]) = 'Без НДС') " & _
                    "AND (([" & tabNameMainBase & "].Количество) Is Null) " & _
                    "AND (([" & tabNameMainBase & "].turnover) = 0) " & _
                "ORDER BY [" & tblNameRegUd1C & "].[" & fldNameRegUd1C & "]"
    Call crtQuery(templateSQL, qryTemplate)
    DoCmd.OpenQuery qryTemplate, acViewNormal
End Function
Function crtUPDATE()
    replaceSQL = "SELECT " & _
                    "REPLACE([registration_number],'/','_') AS TWO, " & _
                    "[" & tblNameRegUd1C & "].pathNum AS ONE " & _
                "FROM [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "];"
    Call crtQuery(replaceSQL, qryUPDATE)
    DoCmd.OpenQuery qryUPDATE, acViewNormal
    DoCmd.Close
    
    Dim dbs As Database
    Set dbs = CurrentDb
    dbs.Execute "UPDATE [" & qryUPDATE & "] " & _
                "SET ONE = TWO " & _
                "WHERE ONE Is Null;"
    dbs.Close
End Function
Function crtGetRu()
' вернуть правую часть строки
', InStrRev([FPath],Chr(92)) AS bS, Len([fpath])-[bS] AS lenFName, Right([fpath],[lenFName]) AS fileName " & _

    getRuSQL = "SELECT " & _
                    "[" & tblNameRegUdResult & "].FPath AS Ссылка, " & _
                    "Код, Наименование, [Ставка НДС] AS [НДС], " & _
                    "registration_number AS [Номер РУ], " & _
                    "[Дата РУ] AS [ДатаРУ], " & _
                    "[Срок действия РУ] AS [срокРУ] " & _
            "FROM ([" & pathMainBase & fileNameLinksToScansRegUd & "].[" & tblNameRegUdResult & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
                    "ON [" & tblNameRegUdResult & "].pathNum = [" & tblNameRegUd1C & "].pathNum) " & _
            "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tblNameRegUd1C & "].registration_number = [" & tabNameMainBase & "].[Номер РУ] " & _
            "WHERE ((([" & tabNameMainBase & "].[Ставка НДС]) " & _
                    "Not Like '20%')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(getRuSQL, qryGetRu)
    DoCmd.OpenQuery qryGetRu, acViewNormal
End Function
Function crtGetDs()
    getDsSQL = "SELECT " & _
                    "endSsDs.ssdsPath AS Ссылка, " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование " & _
                "FROM endSsDs " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endSsDs.ID = [" & tabNameDescr & "].DS " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getDsSQL, qryGetDS)
    DoCmd.OpenQuery qryGetDS, acViewNormal
End Function
Function crtGetSs()
    getSsSQL = "SELECT " & _
                    "endSsDs.ssdsPath AS Ссылка, " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование " & _
                "FROM endSsDs " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endSsDs.ID = [" & tabNameDescr & "].ID " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getSsSQL, qryGetSS)
    DoCmd.OpenQuery qryGetSS, acViewNormal
End Function
Function crtGetUtsi()
    getUtsiSQL = "SELECT " & _
                    "endUtsi.utsiPath AS Ссылка, " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование " & _
                "FROM endUtsi INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                    "ON endUtsi.ID = [" & tabNameDescr & "].UT " & _
                "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(getUtsiSQL, qryGetUtsi)
    DoCmd.OpenQuery qryGetUtsi, acViewNormal
End Function
Function crtExportTo1C()
    exportTo1CSQL = "SELECT [" & tabNameMainBase & "].Код, " & _
                    "[Код группы], " & _
                    "Наименование, " & _
                    "[Наименование для печати], " & _
                    "[Пометка удаления], " & _
                    "[Номер РУ], " & _
                    "[Дата РУ], " & _
                    "[Срок действия РУ], " & _
                    "[Сверено  с РУ/ДС], " & _
                    "[Ставка НДС], " & _
                    "ТНВЭД, " & _
                    "ОКП, " & _
                    "ОКПД2, " & _
                    "поставщикКод, " & _
                    "Производитель, " & _
                    "[Срок действия сертификат соответствия], " & _
                    "Артикул, " & _
                    "[Текстовое описание], " & _
                    "ККМ, " & _
                    "[Срок действия декларация соответствия], [НКМИ], [дата СИ] " & _
            "FROM [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код " & _
            "WHERE ((([" & tabNameDescr & "].примечание) = 'parsing')) " & _
            "ORDER BY Наименование;"
    Call crtQuery(exportTo1CSQL, qryExportTo1C)
    DoCmd.OpenQuery qryExportTo1C, acViewNormal
End Function
Function crtDescr()
    descrSQL = "SELECT " & _
                    "[Код], " & _
                    "примечание, " & _
                    "ID, " & _
                    "DS, " & _
                    "UT " & _
            "FROM [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
            "ORDER BY примечание DESC;"
'    descrSQL = fnDescr_1()
    Call crtQuery(descrSQL, qryDescr)
    DoCmd.OpenQuery qryDescr, acViewNormal
End Function
Function crtDoubleRows()
    doubleRowsSQL = "SELECT " & _
                    "[Код], " & _
                    "Count(Код) AS Повторы " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "GROUP BY Код " & _
            "HAVING (((Count(Код))>1));"
    Call crtQuery(doubleRowsSQL, qryDoubleRows)
    DoCmd.OpenQuery qryDoubleRows, acViewNormal
End Function
Function crtDoubleRowsDel()
    doubleRowsDelSQL = "SELECT [Код], [Пометка удаления] AS [уд] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "WHERE ((([" & tabNameMainBase & "].[Пометка удаления])=2));"
    Call crtQuery(doubleRowsDelSQL, qrydoubleRowsDel)
    DoCmd.OpenQuery qrydoubleRowsDel, acViewNormal
End Function
Function crtFolders()
    foldersSQL = "SELECT " & _
                    "[Код группы], " & _
                    "Группа, " & _
                    "Sum(1) AS Карточек " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "GROUP BY [Код группы], Группа " & _
            "HAVING ((([Код группы]) Is Not Null)) " & _
            "ORDER BY Группа;"
    Call crtQuery(foldersSQL, qryFolders)
'    DoCmd.OpenQuery qryFolders, acViewNormal
End Function
Function crtMainDbEndSsDs()
    mainDBendSsDsSQL = "SELECT [" & tabNameMainBase & "].Код, " & _
                    "[Код группы], " & _
                    "Группа, " & _
                    "Наименование, " & _
                    "[Основной поставщик], " & _
                    "поставщикКод, " & _
                    "[Срок действия сертификат соответствия], " & _
                    "[Срок действия декларация соответствия], " & _
                    "ТНВЭД, " & _
                    "примечание " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
            "WHERE (([" & tabNameMainBase & "].[Группа]) " & _
                        "NOT LIKE '*Материалы*' " & _
                    "AND ([" & tabNameMainBase & "].[Группа]) " & _
                        "NOT LIKE '*Услуги*' " & _
                    "AND ([" & tabNameMainBase & "].[Группа]) " & _
                        "NOT LIKE 'Тара*' " & _
                    "AND NOT ([" & tabNameMainBase & "].[Группа]) " & _
                        "LIKE '999.*' " & _
                    "AND ([" & tabNameMainBase & "].[Группа]) " & _
                        "NOT LIKE '90.*' " & _
                    "AND ([" & tabNameMainBase & "].[turnover]) =1) " & _
            "ORDER BY Наименование;"
    Call crtQuery(mainDBendSsDsSQL, qryMainDBendSsDs)
    DoCmd.OpenQuery qryMainDBendSsDs, acViewNormal
End Function
Function crtMain()
    mainSQL = "SELECT [" & tabNameMainBase & "].Код, " & _
                    "[Код группы] AS КодГр, " & _
                    "Группа, " & _
                    "Наименование, " & _
                    "[Наименование для печати] AS печат, " & _
                    "[Основной поставщик] AS поставщик, " & _
                    "[поставщикКод] AS пствщID, " & _
                    "[Производитель] AS Про, " & _
                    "[Номер РУ] AS РУ, " & _
                    "[Дата РУ] AS ДатРУ, " & _
                    "[Ставка НДС] AS НДС, " & _
                    "ТНВЭД, " & _
                    "ОКП, " & _
                    "ОКПД2, " & _
                    "[" & tabNameMainBase & "].[Срок действия сертификат соответствия] AS [СС], " & _
                    "[" & tabNameMainBase & "].[Срок действия декларация соответствия] AS [ДС], " & _
                    "[Пометка удаления] AS [уд], " & _
                    "[примечание] AS [прим], " & _
                    "[Карточка] AS [созд], " & _
                    "[Артикул] AS [арт], " & _
                    "[turnover] AS [обр] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
            "ORDER BY Наименование;"
    Call crtQuery(mainSQL, qryMain)
    DoCmd.OpenQuery qryMain, acViewNormal
End Function
Function crtMain1()
    main1SQL = "SELECT " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "Группа, " & _
                    "Наименование, " & _
                    "[Наименование для печати] AS [печат], " & _
                    "[Основной поставщик] AS [поставщик], " & _
                    "[Производитель] AS [Про], " & _
                    "[Номер РУ] AS РУ, " & _
                    "[Дата РУ] AS [ДатРУ], " & _
                    "[Ставка НДС] AS [НДС], " & _
                    "ТНВЭД, " & _
                    "ОКП, " & _
                    "ОКПД2, " & _
                    "[Пометка удаления] AS [уд], " & _
                    "[примечание] AS [прим], " & _
                    "[Карточка] AS [созд], " & _
                    "[Артикул] AS [арт], " & _
                    "turnover AS [обр] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
            "WHERE ((([" & tabNameMainBase & "].[Пометка удаления]) = 1)) " & _
            "ORDER BY Наименование;"
    Call crtQuery(main1SQL, qryMain1)
    DoCmd.OpenQuery qryMain1, acViewNormal
End Function
Function crtOpenMainNomenclature()
    mainNomenclatureSQL = "SELECT " & _
                    "[Код], " & _
                    "[Код группы] AS [КодГр], " & _
                    "Группа, " & _
                    "Наименование, " & _
                    "[Основной поставщик] AS [поставщик], [поставщикКод] AS [пствщID], " & _
                    "[Производитель] AS [Про], " & _
                    "[Номер РУ] AS [РУ], " & _
                    "[Дата РУ] AS [ДатРУ], " & _
                    "[Срок действия РУ] AS [endРУ], " & _
                    "[Ставка НДС] AS [НДС], " & _
                    "[Артикул] AS [арт], " & _
                    "[Количество] AS [к-во], " & _
                    "[ЕдХран] AS [хран], " & _
                    "[turnover] AS [обр], " & _
                    "ТНВЭД, ОКП, ОКПД2, " & _
                    "[Срок действия сертификат соответствия] AS [СС], " & _
                    "[Срок действия декларация соответствия] AS [ДС], " & _
                    "[дата СИ] AS [УТСИ], " & _
                    "[Пометка удаления] AS [уд], " & _
                    "[Карточка] AS [созд] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "WHERE ([Наименование] & ' \ ' & [Артикул]) LIKE '*' " & _
            "ORDER BY Наименование;"
    Call crtQuery(mainNomenclatureSQL, qryOpenMainNomenclature)
    DoCmd.OpenQuery qryOpenMainNomenclature, acViewNormal
End Function
Function crtPathRegUd()
    pathRegUdSQL = "SELECT " & _
                    "FPath AS Ссылка " & _
            "FROM [" & pathMainBase & fileNameLinksToScansRegUd & "].RU " & _
            "ORDER BY FPath;"
    Call crtQuery(pathRegUdSQL, qryPathRegUd)
'    DoCmd.OpenQuery qryPathRegUd, acViewNormal
End Function
Function crtPathSertificates()
    pathSertificatesSQL = "SELECT " & _
                    "FPath AS Ссылка " & _
            "FROM [" & pathMainBase & fileNameLinksToScansSertificate & "].RU;"
    Call crtQuery(pathSertificatesSQL, qryPathSertificates)
    DoCmd.OpenQuery qryPathSertificates, acViewNormal
End Function
Function crtRegUdDescr()
    regUdDescrSQL = "SELECT " & _
                    "registration_number, " & _
                    "pathNum " & _
                    "FROM [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "];"
    Call crtQuery(regUdDescrSQL, qryRegUdDescr)
    DoCmd.OpenQuery qryRegUdDescr, acViewNormal
End Function
Function crtProduce()
    produceSQL = "SELECT " & _
                    "Производитель, " & _
                    "Sum(1) AS Выражение1 " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "GROUP BY Производитель " & _
            "HAVING ((Производитель) Is Not Null) " & _
            "ORDER BY Производитель;"
    Call crtQuery(produceSQL, qryProduce)
    DoCmd.OpenQuery qryProduce, acViewNormal
End Function
Function crtRegUdSearch()
'dbRegUd.unique_number,
    regUdSearchSQL = "SELECT " & _
                    "[dbRegUd.registration_number] AS номер, " & _
                    "[dbRegUd.registration_date] AS bgnРУ, " & _
                    "[dbRegUd.registration_date_end] AS endРУ, " & _
                    "[" & tabReestRegUd & "].name, " & _
                    "[" & tabReestRegUd & "].producer, " & _
                    "[" & tabReestRegUd & "].okp, " & _
                    "[" & tabReestRegUd & "].kind " & _
            "FROM [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "];"
    Call crtQuery(regUdSearchSQL, qryRegUdSearch)
    DoCmd.OpenQuery qryRegUdSearch, acViewNormal
End Function
Function crtRegUdSearch1()
    Call crtQuery(regUdSearchSQL, qryRegUdSearch1)
    DoCmd.OpenQuery qryRegUdSearch1, acViewNormal
End Function
Function crtRegUdSearch2()
    Call crtQuery(regUdSearchSQL, qryRegUdSearch2)
    DoCmd.OpenQuery qryRegUdSearch2, acViewNormal
End Function
Function crtSeller()
    sellerSQL = "SELECT " & _
                    "поставщикКод, " & _
                    "[Основной поставщик], " & _
                    "Sum(1) AS total " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "GROUP BY поставщикКод, [Основной поставщик] " & _
            "HAVING (([Основной поставщик]) Is Not Null) " & _
            "ORDER BY [Основной поставщик];"
    Call crtQuery(sellerSQL, qrySeller)
'    DoCmd.OpenQuery qrySeller, acViewNormal
End Function
Function crtSplitRow()
    splitRowSQL = "SELECT " & _
        "SplitRow([Наименование],0) AS 1, " & _
        "SplitRow([Наименование],1) AS 2, " & _
        "SplitRow([Наименование],2) AS 3, " & _
        "SplitRow([Наименование],3) AS 4, " & _
        "SplitRow([Наименование],4) AS 5, " & _
        "SplitRow([Наименование],5) AS 6, " & _
        "SplitRow([Наименование],6) AS 7, " & _
        "SplitRow([Наименование],7) AS 8, " & _
        "SplitRow([Наименование],8) AS 9, " & _
        "[3]&'ZZZ'&[1]&[2]&UCase([5]) AS res " & _
    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
    "WHERE ((([" & tabNameDescr & "].[примечание]) Like 'кУкУ'));"
    Call crtQuery(splitRowSQL, qrySplitRow)
    DoCmd.OpenQuery qrySplitRow, acViewNormal
End Function
Function crtSplitOtkaz()
    splitOtkazSQL = "SELECT " & _
        "SplitOtkaz([FPath],1) AS fileName, " & _
        "[FPath] " & _
    "FROM [" & pathMainBase & fileNameLinksToScansSertificate & "].RU " & _
    "WHERE (((RU.FPath) Like '*отказ *')) " & _
    "ORDER BY SplitOtkaz([FPath],1);"
    Call crtQuery(splitOtkazSQL, qrySplitOtkaz)
'    DoCmd.OpenQuery qrySplitOtkaz, acViewNormal
End Function
Function crtChangesRegUd()
    changesRegUdSQL = "SELECT " & _
                    "Код, " & _
                    "[dataRegUd.registration_number] AS [РУ], " & _
                    "[registration_date] AS [datRU], " & _
                    "[Дата РУ] AS [датРУ], " & _
                    "Left([name],150) AS nameRegUd, " & _
                    "Left([Наименование для печати],150) AS name1C, " & _
                    "IIf([Дата РУ]<>[registration_date],'STOP','') AS [Stop] " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN ([" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
            "INNER JOIN [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "] " & _
                "ON [" & tblNameRegUd1C & "].registration_number = [" & tabReestRegUd & "].registration_number)  " & _
                "ON [" & tabNameMainBase & "].[Номер РУ] = [" & tblNameRegUd1C & "].registration_number " & _
            "GROUP BY " & _
                    "Код, " & _
                    "[" & tblNameRegUd1C & "].registration_number, " & _
                    "registration_date, " & _
                    "[Дата РУ], " & _
                    "Left([name],150), " & _
                    "Left([Наименование для печати],150), " & _
                    "IIf([Дата РУ]<>[registration_date],'STOP','') " & _
            "HAVING ((([" & tblNameRegUd1C & "].registration_number) " & _
                    "Not Like 'ФСЗ 2009/03674' " & _
                    "And ([" & tblNameRegUd1C & "].registration_number) Not Like 'ФСР 2011/09963') AND ((IIf([Дата РУ]<>[registration_date],'STOP',''))='STOP')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(changesRegUdSQL, qryChangesRegUd)
    DoCmd.OpenQuery qryChangesRegUd, acViewNormal
End Function
Function crtOkpRegUd()
    okpRegUdSQL = "SELECT " & _
                    "[" & tabNameMainBase & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование, " & _
                    "[" & tblNameRegUd1C & "].registration_number, " & _
                    "[" & tabReestRegUd & "].registration_date, " & _
                    "[" & tabReestRegUd & "].okp, " & _
                    "[" & tabNameMainBase & "].ОКПД2, " & _
                    "IIf([ОКП]<>[okp],'STOP','') AS Выражение1 " & _
            "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN ([" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
            "INNER JOIN [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "] " & _
                "ON [" & tblNameRegUd1C & "].registration_number = [" & tabReestRegUd & "].registration_number) " & _
                "ON mainNomenclature.[Номер РУ] = [" & tblNameRegUd1C & "].registration_number " & _
            "WHERE (((IIf([ОКП]<>[okp],'STOP',''))='STOP')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(okpRegUdSQL, qryOkpRegUd)
    DoCmd.OpenQuery qryOkpRegUd, acViewNormal
End Function
' !!! что за хрень!!! какой Database1.accdb ??? переделать ! кнопки нет + НЕ РАБОТАЕТ
Function crtChangesRegUdTo1C()
    changesRegUdTo1CSQL = "SELECT " & _
                    "ChgRu.Код, " & _
                    "[" & tblNameRegUd1C & "].registration_number, " & _
                    "Left([name],150) AS baseRegUd, " & _
                    "Left([Наименование для печати],150) AS base1C " & _
            "FROM [" & pathMainBase & "Database1.accdb].ChgRu " & _
            "INNER JOIN ([" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
            "INNER JOIN ([" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
            "INNER JOIN [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "] " & _
                "ON [" & tblNameRegUd1C & "].registration_number = [" & tabReestRegUd & "].registration_number)  " & _
                "ON [" & tabNameMainBase & "].[Номер РУ] = [" & tblNameRegUd1C & "].registration_number)  " & _
                "ON ChgRu.[Код] = [" & tabNameMainBase & "].[" & fldMainCod & "] " & _
            "WHERE ((([" & tblNameRegUd1C & "].registration_number) " & _
                    "Not Like 'ФСЗ 2009/03674' " & _
                    "AND ([" & tblNameRegUd1C & "].registration_number) " & _
                    "Not Like 'ФСР 2011/09963') " & _
                    "AND (([" & tabNameMainBase & "].[Пометка удаления])=1) " & _
                    "AND (([" & tabNameMainBase & "].[Ставка НДС]) Not Like '20%')) " & _
            "ORDER BY [" & tblNameRegUd1C & "].registration_number;"
    Call crtQuery(changesRegUdTo1CSQL, qryChangesRegUdTo1C)
    DoCmd.OpenQuery qryChangesRegUdTo1C, acViewNormal
End Function
Function crtListKeysTableName()
    listKeysTableNameSQL = "SELECT " & _
                    "[" & tblNameRegUd1C & "].pathNum " & _
            "FROM [" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
            "ORDER BY Len([pathNum]) DESC;"
    Call crtQuery(listKeysTableNameSQL, qryListKeysTableName)
    DoCmd.OpenQuery qryListKeysTableName, acViewNormal
End Function
Function crtLinksDoublesRUsearch()
    linksDoublesRUsearchSQL = "SELECT " & _
                    "[" & tblNameRegUdResult & "].pathNum, " & _
                    "Sum(1) AS Выражение1 " & _
            "FROM [" & pathMainBase & fileNameLinksToScansRegUd & "].[" & tblNameRegUdResult & "] " & _
            "GROUP BY [" & tblNameRegUdResult & "].pathNum " & _
            "HAVING ((([" & tblNameRegUdResult & "].pathNum) Is Not Null) " & _
                    "And ((Sum(1)) > 1)) " & _
            "ORDER BY [" & tblNameRegUdResult & "].pathNum;"
    Call crtQuery(linksDoublesRUsearchSQL, qryLinksDoublesRUsearch)
    DoCmd.OpenQuery qryLinksDoublesRUsearch, acViewNormal
End Function
Function crtLinksRUsearch()
    linksRUsearchSQL = "SELECT " & _
                    "[" & tblNameRegUdResult & "].pathNum, " & _
                    "[" & tblNameRegUdResult & "].FPath " & _
            "FROM [" & pathMainBase & fileNameLinksToScansRegUd & "].[" & tblNameRegUdResult & "] " & _
            "ORDER BY [" & tblNameRegUdResult & "].pathNum;"
    Call crtQuery(linksRUsearchSQL, qryLinksRUsearch)
    DoCmd.OpenQuery qryLinksRUsearch, acViewNormal
End Function
Function crtMedCod()
    MedCodSQL = "SELECT " & _
                    "[" & tblNameMedCod_174 & "].Код, " & _
                    "[" & tblNameMedCod_174 & "].[вид классификации медицинских изделий] AS Nomenclature, " & _
                    "[" & tblNameMedCod_174 & "].Наименование AS NameMed " & _
            "FROM [" & pathMainBase & fileNameMedCod_174 & "].[" & tblNameMedCod_174 & "] " & _
            "ORDER BY Наименование;"
    Call crtQuery(MedCodSQL, qryMedCod)
    DoCmd.OpenQuery qryMedCod, acViewNormal
End Function
Function crtMedCod_1c()
    MedCod_1cSQL = "SELECT " & _
                    "[" & tblnameCodIn_1C & "].Код, " & _
                    "[" & tabNameMainBase & "].Наименование, " & _
                    "[" & tabNameMainBase & "].[Ставка НДС] AS [НДС], " & _
                    "[" & tblnameCodIn_1C & "].классф, " & _
                    "[" & tabNameMainBase & "].Количество AS [к-во], " & _
                    "[" & tabNameMainBase & "].turnover AS обр " & _
            "FROM [" & pathMainBase & fileNameMedCod_174 & "].[" & tblnameCodIn_1C & "] " & _
            "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                "ON [" & tblnameCodIn_1C & "].Код = [" & tabNameMainBase & "].Код " & _
            "WHERE (([" & tabNameMainBase & "].Наименование LIKE 'перчатк*' " & _
                "OR [" & tabNameMainBase & "].Наименование LIKE '*комбинез*' " & _
                "OR [" & tabNameMainBase & "].Наименование LIKE 'респиратор*' " & _
                "OR [" & tabNameMainBase & "].Наименование LIKE 'маск*') " & _
                "AND (([" & tabNameMainBase & "].[Пометка удаления])=1))" & _
            "ORDER BY Наименование;"
    Call crtQuery(MedCod_1cSQL, qryMedCodInBase)
    DoCmd.OpenQuery qryMedCodInBase, acViewNormal
End Function
Function crtPP1042()
    getPP1042SQL = "SELECT [" & tblNamePP1042 & "].codID " & _
                "FROM [" & pathMainBase & fileNamePpOkpd2 & "].[" & tblNamePP1042 & "] " & _
                "ORDER BY [" & tblNamePP1042 & "].codID;"
    Call crtQuery(getPP1042SQL, qryPP1042)
    DoCmd.OpenQuery qryPP1042, acViewNormal
End Function
Function crtPP688()
    getPP688SQL = "SELECT [" & tblNamePP688 & "].codID " & _
                "FROM [" & pathMainBase & fileNamePpOkpd2 & "].[" & tblNamePP688 & "] " & _
                "ORDER BY [" & tblNamePP688 & "].codID;"
    Call crtQuery(getPP688SQL, qryPP688)
    DoCmd.OpenQuery qryPP688, acViewNormal
End Function
Function crtFTSpp312()
    getFTSpp312SQL = "SELECT [" & tblNameFTSpp312 & "].tnved,[" & tblNameFTSpp312 & "].descrT4,[" & tblNameFTSpp312 & "].dateRow,[" & tblNameFTSpp312 & "].dateOut,[" & tblNameFTSpp312 & "].prim " & _
                "FROM [" & pathMainBase & fileNameFTSpp312 & "].[" & tblNameFTSpp312 & "] " & _
                "ORDER BY [" & tblNameFTSpp312 & "].tnved;"
    Call crtQuery(getFTSpp312SQL, qryFTSpp312)
    DoCmd.OpenQuery qryFTSpp312, acViewNormal
End Function
Function crtKaNomenclature()
    kaNomenclatureSQL = "SELECT " & _
                    "[" & fldKaCod & "], " & _
                    "[" & fldKaName & "], " & _
                    "[" & tabNameKaBase & "].[Ставка НДС] AS [НДС] " & _
            "FROM [" & pathMainBase & fileNameKaBase & "].[" & tabNameKaBase & "] " & _
            "ORDER BY Наименование;"
    Call crtQuery(kaNomenclatureSQL, qryKaNomenclature)
    DoCmd.OpenQuery qryKaNomenclature, acViewNormal
End Function
Function crtKaJoinMain()
    kaJoinMainSQL = "SELECT " & _
                    "[" & tabNameKaBase & "].[" & fldKaCod & "] AS [KA], " & _
                    "[" & tabNameKaBase & "].[" & fldKaName & "] AS [KA NAME], " & _
                    "[" & tabNameMainBase & "].[" & fldMainCod & "] AS [YT]" & _
                "FROM [" & pathMainBase & fileNameKaBase & "].[" & tabNameKaBase & "] " & _
                "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameKaBase & "].[" & fldKaCod & "] = [" & tabNameMainBase & "].[" & fldMainCod & "] " & _
                "ORDER BY [" & tabNameKaBase & "].[" & fldKaName & "];"
    Call crtQuery(kaJoinMainSQL, qryKaJoinMain)
    DoCmd.OpenQuery qryKaJoinMain, acViewNormal
End Function
Function crtKaRightMain()
    kaRightMainSQL = "SELECT " & _
                    "[" & tabNameKaBase & "].[" & fldKaCod & "] AS [KA], " & _
                    "[" & tabNameKaBase & "].[" & fldKaName & "] AS [KA NAME], " & _
                    "[" & tabNameMainBase & "].[" & fldMainCod & "] AS [YT], " & _
                    "[" & tabNameMainBase & "].[" & fldMainName & "] AS [YT NAME]" & _
                "FROM [" & pathMainBase & fileNameKaBase & "].[" & tabNameKaBase & "] " & _
                "RIGHT JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "ON [" & tabNameKaBase & "].[" & fldKaCod & "] = [" & tabNameMainBase & "].[" & fldMainCod & "] " & _
                "ORDER BY [" & tabNameMainBase & "].[" & fldMainName & "];"
    Call crtQuery(kaRightMainSQL, qryKaRightMain)
    DoCmd.OpenQuery qryKaRightMain, acViewNormal
End Function



