Attribute VB_Name = "mdlCodSQLendSsDs"
Option Compare Database
Option Explicit
'��� ������: "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, "
'���������� ��� ���� �� ������� ������� ??? ������, ����� ������� � ������� Excel
Function crtEndSsTo1c()
    endSsTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[������������ ��� ������], " & _
                        "[" & tabNameMainBase & "].[������� ��������], " & _
                        "[" & tabNameMainBase & "].[����� ��], " & _
                        "[" & tabNameMainBase & "].[���� ��], " & _
                        "[" & tabNameMainBase & "].[���� �������� ��], " & _
                        "[" & tabNameMainBase & "].[�������  � ��/��], " & _
                        "[" & tabNameMainBase & "].[������ ���], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].����2, " & _
                                "[" & tabNameMainBase & "].������������, [" & tabNameMainBase & "].�������������, " & _
                                    "[dd] & '.' & [mm] & '.' & [yy] AS [���� �������� ���������� ������������], " & _
                                        "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, " & _
                                            "[" & tabNameMainBase & "].�������, [" & tabNameMainBase & "].[��������� ��������], [" & tabNameMainBase & "].���, [" & tabNameMainBase & "].[���� �������� ���������� ������������], [����], [���� ��] " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].��� = [" & tabNameMainBase & "].���) " & _
                        "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].ID " & _
                    "WHERE ((([" & tblNameSsDs & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
    Call crtQuery(endSsTo1cSQL, qryEndSsTo1C)
    DoCmd.OpenQuery qryEndSsTo1C, acViewNormal
End Function
Function crtEndDsTo1c()
    endDsTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[������������ ��� ������], " & _
                        "[" & tabNameMainBase & "].[������� ��������], " & _
                        "[" & tabNameMainBase & "].[����� ��], " & _
                        "[" & tabNameMainBase & "].[���� ��], " & _
                        "[" & tabNameMainBase & "].[���� �������� ��], " & _
                        "[" & tabNameMainBase & "].[�������  � ��/��], " & _
                        "[" & tabNameMainBase & "].[������ ���], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].����2, " & _
                                "[" & tabNameMainBase & "].������������, [" & tabNameMainBase & "].�������������, [" & tabNameMainBase & "].[���� �������� ���������� ������������], " & _
                                    "[" & tabNameMainBase & "].�������, [" & tabNameMainBase & "].[��������� ��������], [" & tabNameMainBase & "].���, " & _
                                        "[dd] & '.' & [mm] & '.' & [yy] AS [���� �������� ���������� ������������], " & _
                                            "Mid([" & tblNameSsDs & "],8,2) AS dd, Mid([" & tblNameSsDs & "],6,2) AS mm, 20 & Mid([" & tblNameSsDs & "],4,2) AS yy, [����], [���� ��] " & _
                    "FROM [" & pathMainBase & fileNameSsDs & "].[" & tblNameSsDs & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].��� = [" & tabNameMainBase & "].���) " & _
                        "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].DS " & _
                    "WHERE ((([" & tblNameSsDs & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
    Call crtQuery(endDsTo1cSQL, qryEndDsTo1C)
    DoCmd.OpenQuery qryEndDsTo1C, acViewNormal
End Function
Function crtEndUtsiTo1C()
    endUtsiTo1cSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[������������ ��� ������], " & _
                        "[" & tabNameMainBase & "].[������� ��������], " & _
                        "[" & tabNameMainBase & "].[����� ��], " & _
                        "[" & tabNameMainBase & "].[���� ��], " & _
                        "[" & tabNameMainBase & "].[���� �������� ��], " & _
                        "[" & tabNameMainBase & "].[�������  � ��/��], " & _
                        "[" & tabNameMainBase & "].[������ ���], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].����2, " & _
                            "[" & tabNameMainBase & "].������������, [" & tabNameMainBase & "].�������������, [���� �������� ���������� ������������], [" & tabNameMainBase & "].�������, " & _
                                "[" & tabNameMainBase & "].[��������� ��������], [" & tabNameMainBase & "].���, [" & tabNameMainBase & "].[���� �������� ���������� ������������], [����], " & _
                                    "[dd] & '.' & [mm] & '.' & [yy] AS [���� ��], " & _
                                        "Mid(([" & tblNameUtsi & "].utsi),10,2) AS dd, Mid(([" & tblNameUtsi & "].utsi),8,2) AS mm, 20 & Mid(([" & tblNameUtsi & "].utsi),6,2) AS yy " & _
                    "FROM [" & pathMainBase & fileNameUtsi & "].[" & tblNameUtsi & "] " & _
                    "INNER JOIN ([" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                        "ON [" & tabNameDescr & "].��� = [" & tabNameMainBase & "].���) " & _
                        "ON [" & tblNameUtsi & "].ID = [" & tabNameDescr & "].UT " & _
                    "WHERE ((([" & tblNameUtsi & "].X) = 'u')) " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
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
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[�������� ���������]," & _
                        "[" & tabNameMainBase & "].����������, " & _
                        "[" & tabNameDescr & "].DS, " & _
                        "[" & tabNameDescr & "].ID, " & _
                        "[" & tabNameDescr & "].UT " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].[" & fieldName & "] = [" & tabNameDescr & "].[" & fieldName & "] " & _
                    "WHERE ((([" & tabNameMainBase & "].������������) " & _
                            "Not Like '*����*' " & _
                        "And ([" & tabNameMainBase & "].������������) " & _
                            "Not Like '*�����*' " & _
                        "And ([" & tabNameMainBase & "].������������) " & _
                            "Not Like '*�������*') " & _
                        "AND ((mainNomenclature.[�������� ���������]) Is Null) " & _
                        "AND ((mainNomenclature.����������) Is Not Null) " & _
                        "AND (([DS]+[ID]+[UT])=0))" & _
                    "ORDER BY [" & tabNameMainBase & "].���������� DESC"
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
                        "InStr(1,[FPath],'\��') AS symb, " & _
                        "Len([FPath])-[symb] AS b, " & _
                        "[" & strTableName & "].FPath, " & _
                        "Right([FPath],[b]) AS res " & _
                    "FROM [" & pathMainBase & fileNameLinksToScansSertificate & "].[" & strTableName & "] " & _
                    "WHERE (((InStr(1,[FPath],'\��'))>0))"
    Call crtQuery(fromLinksSQL, qryFromLinksToScansSertificates)
    DoCmd.OpenQuery qryFromLinksToScansSertificates, acViewNormal
End Function
Function crtMainDS()
    mainDsSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������, " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[�������� ���������], " & _
                        "[" & tabNameMainBase & "].������������ AS [��������], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].turnover AS [���], " & _
                        "[" & tabNameMainBase & "].���������� AS [�����], " & _
                        "[" & tabNameDescr & "].DS " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].��� = [" & tabNameDescr & "].��� " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
    Call crtQuery(mainDsSQL, qryMainDs)
    DoCmd.OpenQuery qryMainDs, acViewNormal
End Function
Function crtMainSS()
    mainSsSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������, " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[�������� ���������], " & _
                        "[" & tabNameMainBase & "].������������ AS [��������], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].turnover AS [���], " & _
                        "[" & tabNameMainBase & "].���������� AS [�����], " & _
                        "[" & tabNameDescr & "].ID " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].��� = [" & tabNameDescr & "].��� " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
    Call crtQuery(mainSsSQL, qryMainSs)
    DoCmd.OpenQuery qryMainSs, acViewNormal
End Function
Function crtMainUTSI()
    mainUtsiSQL = "SELECT " & _
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].[��� ������], " & _
                        "[" & tabNameMainBase & "].������, " & _
                        "[" & tabNameMainBase & "].������������, " & _
                        "[" & tabNameMainBase & "].[�������� ���������], " & _
                        "[" & tabNameMainBase & "].������������ AS [��������], " & _
                        "[" & tabNameMainBase & "].�����, " & _
                        "[" & tabNameMainBase & "].turnover AS [���], " & _
                        "[" & tabNameMainBase & "].���������� AS [�����], " & _
                        "[" & tabNameDescr & "].UT " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "INNER JOIN [" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] " & _
                        "ON [" & tabNameMainBase & "].��� = [" & tabNameDescr & "].��� " & _
                    "ORDER BY [" & tabNameMainBase & "].������������"
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
                        "[" & tabNameMainBase & "].���, " & _
                        "[" & tabNameMainBase & "].������������, " & _
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
                            "ON [" & tblNameFirm & "].[������������] = [" & tblNamePost & "].[������������]) " & _
                            "ON [" & tabNameMainBase & "].[������������] = [" & tblNameFirm & "].[������������]) " & _
                            "ON [" & tabNameDescr & "].��� = [" & tabNameMainBase & "].���) " & _
                            "ON [" & tblNameSsDs & "].ID = [" & tabNameDescr & "].DS " & _
                        "WHERE ((([" & tabNameMainBase & "].turnover)=1) " & _
                            "AND (([" & tblNameSsDs & "].X)='a')) " & _
                        "ORDER BY [" & tabNameDescr & "].DS"
    Call crtQuery(SsDsPostDsSQL, qrySsDsPostDs)
    DoCmd.OpenQuery qrySsDsPostDs, acViewNormal
End Function
