Attribute VB_Name = "mdlEditMainBase"
Option Compare Database
Option Explicit

Function editMainBase(dbPathAndName As String)
    Dim accApp
    Set accApp = New Access.Application
    accApp.OpenCurrentDatabase dbPathAndName, acNewDatabaseFormatUserDefault
    crtMainBaseSQLDelRow
    
    Dim dbs As Database
    Set dbs = OpenDatabase(pathMainBase & fileNameMainBase)
'разборка полей содержащих ДАТЫ
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Карточка] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Дата РУ] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Срок действия РУ] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Срок действия сертификат соответствия] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Срок действия декларация соответствия] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [дата СИ] DATE"
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Счета] DATE"
    
    dbs.Execute "UPDATE mainNomenclature SET [Карточка] = LEFT([txtDCard],10);"
    dbs.Execute "UPDATE mainNomenclature SET [Дата РУ] = LEFT([txtDBgnRegU],10);"
    dbs.Execute "UPDATE mainNomenclature SET [Срок действия РУ] = LEFT([txtDEndRegU],10);"
    dbs.Execute "UPDATE mainNomenclature SET [Срок действия сертификат соответствия] = LEFT([txtDSs],10);"
    dbs.Execute "UPDATE mainNomenclature SET [Срок действия декларация соответствия] = LEFT([txtDDs],10);"
    dbs.Execute "UPDATE mainNomenclature SET [дата СИ] = LEFT([txtUTSI],10);"
    dbs.Execute "UPDATE mainNomenclature SET [Счета] = LEFT([txtDBill],10);"
' Да/Нет
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [Пометка удаления] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [Пометка удаления] = 1 WHERE [txtDel] = 'Нет';"
    dbs.Execute "UPDATE mainNomenclature SET [Пометка удаления] = 2 WHERE [txtDel] = 'Да';"
    
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [turnover] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [turnover] = 0 WHERE [txtTurnover] = 'Нет';"
    dbs.Execute "UPDATE mainNomenclature SET [turnover] = 1 WHERE [txtTurnover] = 'Да';"
    
    dbs.Execute "ALTER TABLE mainNomenclature ADD COLUMN [trace] INTEGER"
    dbs.Execute "UPDATE mainNomenclature SET [trace] = 0 WHERE [txtTrace] = 'Нет';"
    dbs.Execute "UPDATE mainNomenclature SET [trace] = 1 WHERE [txtTrace] = 'Да';"

'    dbs.Execute "ALTER TABLE mainNomenclature DROP COLUMN textDate DATE"
'    dbs.Execute "ALTER TABLE mainNomenclature ALTER COLUMN [Код] TEXT(11)"
    
    Set dbs = Nothing
'необходимо прокрутить запросы для обновления в них пути
    crtDoubleRows
    DoCmd.Close
    crtDoubleRowsDel
    DoCmd.Close

    fncDelDoubleRow 'удаление дублей карточек
    crtMainBaseSQLtoDescr 'добавление новых кодов карточек в Descr.accdb
    crtMainBaseSQLtoRegUd1C 'добавление "Номер РУ" в dbRegUd1C.accdb
    accApp.Quit
End Function
Sub editNames()
    Dim dbs As Database
    Set dbs = CurrentDb()

    dbs.Execute "UPDATE main SET [Наименование] = REPLACE([Наименование],Chr(10),'');" 'Linefeed character символ перевода строки
    dbs.Execute "UPDATE main SET [печат] = REPLACE([печат],Chr(10),'');" 'Linefeed character символ перевода строки
'    dbs.Execute "UPDATE main SET [ККМ] = REPLACE([ККМ],Chr(10),'');" 'это поле НЕ ВЫВЕДЕНО в таблицу
End Sub

