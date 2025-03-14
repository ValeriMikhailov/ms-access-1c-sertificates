Attribute VB_Name = "mdlCodSQLregUdBase"
Option Compare Database
Option Explicit

Function crtRegUdSQLdelRow()
    regUdSQLdelRow = "DELETE " & _
                        "[" & tabReestRegUd & "].[registration_number] " & _
                    "FROM [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "] " & _
                    "WHERE [registration_number]='a'"
    Call crtQuery(regUdSQLdelRow, qryRegUdDelRow)
    DoCmd.OpenQuery qryRegUdDelRow, acViewNormal
End Function
Function crtMainBaseSQLtoDescr()
    mainBaseSQLtoDescr = "INSERT INTO " & _
                    "[" & pathMainBase & fileNameDescr & "].[" & tabNameDescr & "] (" & fieldName & ") " & _
                    "SELECT " & _
                    "[" & tabNameMainBase & "].[" & fieldName & "] " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "];"
    Call crtQuery(mainBaseSQLtoDescr, qryMainBaseToDescr)
    DoCmd.OpenQuery qryMainBaseToDescr, acViewNormal
End Function
Function crtMainBaseSQLtoRegUd1C()
    mainBaseSQLtoRegUd1C = "INSERT INTO " & _
                    "[" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] (" & fldNameRegUd1C & ") " & _
                    "SELECT " & _
                    "[" & tabNameMainBase & "].[" & fldNameNumRegUd & "] " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "];"
    Call crtQuery(mainBaseSQLtoRegUd1C, qryMainBaseToRegUd1C)
    DoCmd.OpenQuery qryMainBaseToRegUd1C, acViewNormal
End Function
Function crtMainBaseSQLDelRow()
    mainBaseSQLDelRow = "DELETE " & _
                        "[" & tabNameMainBase & "].[" & fieldName & "] " & _
                    "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                    "WHERE [" & fieldName & "]='a'"
    Call crtQuery(mainBaseSQLDelRow, qryMainBaseDelRow)
    DoCmd.OpenQuery qryMainBaseDelRow, acViewNormal
End Function
Function crtDateEndRegUd()
    dateEndRegUdSQL = "SELECT " & _
                    "[" & tblNameRegUd1C & "].registration_number AS RegNum, " & _
                    "[Срок действия РУ] AS ru1C, " & _
                    "registration_date_end AS ruReestr, " & _
                    "IIF ([Срок действия РУ]=[registration_date_end],'ok','STOP') AS check, " & _
                    "[Наименование], " & _
                    "[Код] " & _
                "FROM [" & pathMainBase & fileNameMainBase & "].[" & tabNameMainBase & "] " & _
                "INNER JOIN ([" & pathMainBase & fileNameRegUd1C & "].[" & tblNameRegUd1C & "] " & _
                "INNER JOIN [" & pathRegUdBase & fileNameRegUdBase & "].[" & tabReestRegUd & "] " & _
                    "ON [" & tblNameRegUd1C & "].registration_number = [" & tabReestRegUd & "].registration_number) " & _
                    "ON [" & tabNameMainBase & "].[Номер РУ] = [" & tblNameRegUd1C & "].registration_number " & _
                "WHERE ((([" & tblNameRegUd1C & "].registration_number) " & _
                        "Not Like 'ФСЗ 2009/03674' " & _
                    "And ([" & tblNameRegUd1C & "].registration_number) " & _
                        "Not Like 'ФСР 2011/09963') " & _
                    "AND ((dbRegUd.registration_date_end) " & _
                        "Is Not Null) " & _
                    "AND (([" & tabNameMainBase & "].[Ставка НДС]) " & _
                        "Not Like '20%')) " & _
                "ORDER BY [" & tabReestRegUd & "].registration_date_end DESC, [" & tblNameRegUd1C & "].registration_number"
    Call crtQuery(dateEndRegUdSQL, qryDateEndregUd)
    DoCmd.OpenQuery qryDateEndregUd, acViewNormal
End Function
