Attribute VB_Name = "mdlCodSQLendSsDs"
Option Compare Database
Option Explicit
'как скрыть: "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, "
'возвращать эти поля из другого запроса ??? скорее, лучше удалять в макросе Excel
Function crtEndSsTo1c()
    endSsTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Наименование для печати], " & _
                        "[" & tabNameMainBase & "].[Пометка удаления], " & _
                        "[" & tabNameMainBase & "].[Номер РУ], " & _
                        "[" & tabNameMainBase & "].[Дата РУ], " & _
                        "[" & tabNameMainBase & "].[Срок действия РУ], " & _
                        "[" & tabNameMainBase & "].[Сверено  с РУ/ДС], " & _
                        "[" & tabNameMainBase & "].[Ставка НДС], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].ОКП, " & _
                        "[" & tabNameMainBase & "].ОКПД2, " & _
                                "[" & tabNameMainBase & "].поставщикКод, [" & tabNameMainBase & "].Производитель, " & _
                                    "[dd] & '.' & [mm] & '.' & [yy] AS [Срок действия сертификат соответствия], " & _
                                        "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, " & _
                                            "[" & tabNameMainBase & "].Артикул, [" & tabNameMainBase & "].[Текстовое описание], [" & tabNameMainBase & "].ККМ, [" & tabNameMainBase & "].[Срок действия декларация соответствия], [НКМИ], [дата СИ] " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                        "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].ID " & _
                    "WHERE ((([" & tblNameSsDs & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(endSsTo1cSQL, qryEndSsTo1C)
    DoCmd.OpenQuery qryEndSsTo1C, acViewNormal
End Function
Function crtEndDsTo1c()
    endDsTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Наименование для печати], " & _
                        "[" & tabNameMainBase & "].[Пометка удаления], " & _
                        "[" & tabNameMainBase & "].[Номер РУ], " & _
                        "[" & tabNameMainBase & "].[Дата РУ], " & _
                        "[" & tabNameMainBase & "].[Срок действия РУ], " & _
                        "[" & tabNameMainBase & "].[Сверено  с РУ/ДС], " & _
                        "[" & tabNameMainBase & "].[Ставка НДС], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].ОКП, " & _
                        "[" & tabNameMainBase & "].ОКПД2, " & _
                                "[" & tabNameMainBase & "].поставщикКод, [" & tabNameMainBase & "].Производитель, [" & tabNameMainBase & "].[Срок действия сертификат соответствия], " & _
                                    "[" & tabNameMainBase & "].Артикул, [" & tabNameMainBase & "].[Текстовое описание], [" & tabNameMainBase & "].ККМ, " & _
                                        "[dd] & '.' & [mm] & '.' & [yy] AS [Срок действия декларация соответствия], " & _
                                            "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, [НКМИ], [дата СИ] " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                        "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].DS " & _
                    "WHERE ((([" & tblNameSsDs & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(endDsTo1cSQL, qryEndDsTo1C)
    DoCmd.OpenQuery qryEndDsTo1C, acViewNormal
End Function
Function crtEndUtsiTo1C()
    endUtsiTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Наименование для печати], " & _
                        "[" & tabNameMainBase & "].[Пометка удаления], " & _
                        "[" & tabNameMainBase & "].[Номер РУ], " & _
                        "[" & tabNameMainBase & "].[Дата РУ], " & _
                        "[" & tabNameMainBase & "].[Срок действия РУ], " & _
                        "[" & tabNameMainBase & "].[Сверено  с РУ/ДС], " & _
                        "[" & tabNameMainBase & "].[Ставка НДС], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].ОКП, " & _
                        "[" & tabNameMainBase & "].ОКПД2, " & _
                            "[" & tabNameMainBase & "].поставщикКод, [" & tabNameMainBase & "].Производитель, [Срок действия сертификат соответствия], [" & tabNameMainBase & "].Артикул, " & _
                                "[" & tabNameMainBase & "].[Текстовое описание], [" & tabNameMainBase & "].ККМ, [" & tabNameMainBase & "].[Срок действия декларация соответствия], [НКМИ], " & _
                                    "[dd] & '.' & [mm] & '.' & [yy] AS [дата СИ], " & _
                                        "Mid(([" & tblNameUtsi & "].utsi),10,2) AS dd, Mid(([" & tblNameUtsi & "].utsi),8,2) AS mm, 20 & Mid(([" & tblNameUtsi & "].utsi),6,2) AS yy " & _
                    "FROM [" & pathMainBase & fileNameUtsi & "].[" & tblNameUtsi & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                        "ON [" & tblNameUtsi & "].ID = [" & tabNameDescr & "].UT " & _
                    "WHERE ((([" & tblNameUtsi & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(endUtsiTo1cSQL, qryEndUtsiTo1C)
    DoCmd.OpenQuery qryEndUtsiTo1C, acViewNormal
End Function
Function crtEndSsDs()
    endSsDsSQL = "SELECT " & _
                        "[" & tblNameSsDs & "].ID, " & _
                        "[" & tblNameSsDs & "].X, " & _
                        "[" & tblNameSsDs & "].ssds, " & _
                        "[" & tblNameSsDs & "].ssdsPath, " & _
                        "Mid([" & tblNameSsDs & "],8,2) AS dd, " & _
                        "Mid([" & tblNameSsDs & "],6,2) AS mm, " & _
                        "20 & Mid([" & tblNameSsDs & "],4,2) AS yy, " & _
                        "[dd] & '.' & [mm] & '.' & [yy] AS endDate " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "ORDER BY [" & tblNameSsDs & "].X"
    Call crtQuery(endSsDsSQL, qryEndSsDs)
    DoCmd.OpenQuery qryEndSsDs, acViewNormal
End Function
Function crtEndUtsi()
    endUtsiSQL = "SELECT " & _
                        "[" & tblNameUtsi & "].ID, " & _
                        "[" & tblNameUtsi & "].X, " & _
                        "[" & tblNameUtsi & "].utsi, " & _
                        "[" & tblNameUtsi & "].utsiPath, " & _
                        "Mid([utsi],10,2) AS dd, " & _
                        "Mid([utsi],8,2) AS mm, " & _
                        "20 & Mid([utsi],6,2) AS yy, " & _
                        "[dd] & '.' & [mm] & '.' & [yy] AS endDate " & _
                    "FROM [" & pathMainBase & fileNameUtsi & "].[" & tblNameUtsi & "] " & _
                    "WHERE ((([" & tblNameUtsi & "].X) " & _
                        "Not Like 'z'))" & _
                    "ORDER BY [" & tblNameUtsi & "].utsi"
    Call crtQuery(endUtsiSQL, qryEndUtsi)
    DoCmd.OpenQuery qryEndUtsi, acViewNormal
End Function
Function crtEndDsSsLeftover()
    endDsSsLeftoverSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].[" & fieldName & "], " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Основной поставщик]," & _
                        "[" & tabNameMainBase & "].Количество, " & _
                        "[" & tabNameDescr & "].DS, " & _
                        "[" & tabNameDescr & "].ID, " & _
                        "[" & tabNameDescr & "].UT " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].[" & fieldName & "] = [" & tabNameDescr & "].[" & fieldName & "] " & _
                    "WHERE ((([" & tabNameMainBase & "].Наименование) " & _
                            "Not Like '*труд*' " & _
                        "And ([" & tabNameMainBase & "].Наименование) " & _
                            "Not Like '*уценк*' " & _
                        "And ([" & tabNameMainBase & "].Наименование) " & _
                            "Not Like '*спецком*') " & _
                        "AND ((mainNomenclature.[Основной поставщик]) Is Null) " & _
                        "AND ((mainNomenclature.Количество) Is Not Null) " & _
                        "AND (([DS]+[ID]+[UT])=0))" & _
                    "ORDER BY [" & tabNameMainBase & "].Количество DESC"
    Call crtQuery(endDsSsLeftoverSQL, qryEndDsSsLeftover)
    DoCmd.OpenQuery qryEndDsSsLeftover, acViewNormal
End Function
Function crtToSsDs()
    toSsDsSQL = "SELECT " & _
                        "[" & tblNameSsDs & "].ssdsPath, " & _
                        "[" & tblNameSsDs & "].ssds " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "ORDER BY [" & tblNameSsDs & "].X"
    Call crtQuery(toSsDsSQL, qryToSsDs)
    DoCmd.OpenQuery qryToSsDs, acViewNormal
End Function
Function crtToUTSI()
    toUtsiSQL = "SELECT " & _
                        "[" & tblNameUtsi & "].utsiPath, " & _
                        "[" & tblNameUtsi & "].utsi " & _
                    "FROM [" & pathMainBase & fileNameUtsi & "].[" & tblNameUtsi & "] " & _
                    "ORDER BY [" & tblNameUtsi & "].X"
    Call crtQuery(toUtsiSQL, qryToUtsi)
    DoCmd.OpenQuery qryToUtsi, acViewNormal
End Function
Function crtFromLinksToScansSertificates()
    fromLinksSQL = "SELECT " & _
                        "InStr(1,[FPath],'\дс') AS symb, " & _
                        "Len([FPath])-[symb] AS b, " & _
                        "[" & strTableName & "].FPath, " & _
                        "Right([FPath],[b]) AS res " & _
                    "FROM [" & pathMainBase & fileNameLinksToScansSertificate & "].[" & strTableName & "] " & _
                    "WHERE (((InStr(1,[FPath],'\дс'))>0))"
    Call crtQuery(fromLinksSQL, qryFromLinksToScansSertificates)
    DoCmd.OpenQuery qryFromLinksToScansSertificates, acViewNormal
End Function
Function crtMainDS()
    mainDsSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Группа, " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Основной поставщик], " & _
                        "[" & tabNameMainBase & "].поставщикКод AS [пствщКод], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].turnover AS [обр], " & _
                        "[" & tabNameMainBase & "].Количество AS [колВо], " & _
                        "[" & tabNameDescr & "].DS " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(mainDsSQL, qryMainDs)
    DoCmd.OpenQuery qryMainDs, acViewNormal
End Function
Function crtMainSS()
    mainSsSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Группа, " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Основной поставщик], " & _
                        "[" & tabNameMainBase & "].поставщикКод AS [пствщКод], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].turnover AS [обр], " & _
                        "[" & tabNameMainBase & "].Количество AS [колВо], " & _
                        "[" & tabNameDescr & "].ID " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(mainSsSQL, qryMainSs)
    DoCmd.OpenQuery qryMainSs, acViewNormal
End Function
Function crtMainUTSI()
    mainUtsiSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].[Код группы], " & _
                        "[" & tabNameMainBase & "].Группа, " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tabNameMainBase & "].[Основной поставщик], " & _
                        "[" & tabNameMainBase & "].поставщикКод AS [пствщКод], " & _
                        "[" & tabNameMainBase & "].ТНВЭД, " & _
                        "[" & tabNameMainBase & "].turnover AS [обр], " & _
                        "[" & tabNameMainBase & "].Количество AS [колВо], " & _
                        "[" & tabNameDescr & "].UT " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].Код = [" & tabNameDescr & "].Код " & _
                    "ORDER BY [" & tabNameMainBase & "].Наименование"
    Call crtQuery(mainUtsiSQL, qryMainUtsi)
    DoCmd.OpenQuery qryMainUtsi, acViewNormal
End Function
Function crtSsDsMovies()
    SsDsMoviesSQL = "SELECT " & _
                        "[" & tblNameSsDs & "].ID, " & _
                        "[" & tblNameSsDs & "].X, " & _
                        "Mid([" & tblNameSsDs & "],4,6) AS d, " & _
                        "Len([" & tblNameSsDs & "])-10 AS b, " & _
                        "Right([" & tblNameSsDs & "],[b]) AS Expr3, " & _
                        "[" & tblNameSsDs & "].ssdsPath " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "WHERE ((([" & tblNameSsDs & "].X) " & _
                            "Not Like 'z') " & _
                        "AND ([" & tblNameSsDs & "].X) " & _
                            "Not Like 'y') " & _
                    "ORDER BY Mid([" & tblNameSsDs & "],4,6);"
    Call crtQuery(SsDsMoviesSQL, qrySsDsMovies)
    DoCmd.OpenQuery qrySsDsMovies, acViewNormal
End Function
Function crtSsDsPostDs()
    SsDsPostDsSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].Код, " & _
                        "[" & tabNameMainBase & "].Наименование, " & _
                        "[" & tblNameFirm & "].partner, " & _
                        "[" & tblNamePost & "].contact, " & _
                            "[" & tblNamePost & "].email, " & _
                        "[" & tblNamePost & "].phone, " & _
                        "[" & tabNameDescr & "].DS " & _
                        "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                        "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "INNER JOIN ([" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "INNER JOIN ([" & pathMainBase & fileNamePost & "].[" & tblNameFirm & "] " & _
                        "INNER JOIN [" & pathMainBase & fileNamePost & "].[" & tblNamePost & "] " & _
                            "ON [" & tblNameFirm & "].[поставщикКод] = [" & tblNamePost & "].[поставщикКод]) " & _
                            "ON [" & tabNameMainBase & "].[поставщикКод] = [" & tblNameFirm & "].[поставщикКод]) " & _
                            "ON [" & tabNameDescr & "].Код = [" & tabNameMainBase & "].Код) " & _
                            "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].DS " & _
                        "WHERE ((([" & tabNameMainBase & "].turnover)=1) " & _
                            "AND (([" & tblNameSsDs & "].X)='a')) " & _
                        "ORDER BY [" & tabNameDescr & "].DS"
    Call crtQuery(SsDsPostDsSQL, qrySsDsPostDs)
    DoCmd.OpenQuery qrySsDsPostDs, acViewNormal
End Function
