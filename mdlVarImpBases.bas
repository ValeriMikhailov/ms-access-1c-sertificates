Attribute VB_Name = "mdlVarImpBases"
Option Compare Database
Option Explicit
' ------------ Excel macroses ----------------
    Public Const fileExcelMacro As String = "macro.xlsm"
    Public Const nameExcelMacroImpMainBase As String = "impShtMainBase"
    Public Const nameExcelMacroImpRegUd As String = "importRegUd"
    Public Const nameExcelMacroAuthor As String = "author"
    Public Const nameExcelMacroReceipts As String = "receipts"
' ------------ import Mainbase -------------------
    Public Const fileNameMainBase As String = "mainBase.accdb"
    Public Const fileExcelMainBase As String = "upLoad.xlsx"
    Public Const tabNameMainBase As String = "mainNomenclature"
' ------------ editing with import Mainbase ----------------
    Public Const fileNameRegUd1C As String = "dbRegUd1C.accdb"
    Public Const tblNameRegUd1C As String = "dataRegUd"
    Public Const fldNameRegUd1C As String = "registration_number"
    Public Const fldNameNumRegUd As String = "Номер РУ"
' ------------ import Database RegUd ----------------
    Public Const pathRegUdBase As String = "D:\reestrRegUd\"
    Public Const fileNameRegUdBase As String = "reestrRegUd.accdb"
    Public Const fileExcelReestRegUd As String = "regUd.xlsx"
    Public Const tabReestRegUd As String = "dbRegUd"
' ------------ import Receipts ----------------
    Public Const fileNameReceipts As String = "receipts.accdb"
    Public Const fileExcelReceipts As String = "receipts.xlsx"
    Public Const tabReceipts As String = "receipts"
    Public Const frmnameReceiptsRegUd As String = "receiptsRegUd"
' ------------ import Author ----------------
    Public Const fileNameAuthor As String = "author.accdb"
    Public Const fileExcelAuthor As String = "author.xlsx"
    Public Const tabAuthor As String = "author"
    Public Const frmNameDatesInsert As String = "DatesInsert"
    Public SD, FD As String
'    Public DateEnd As String ' Form DateInsert
'    Public DateBegin As String ' Form DateInsert
' ------------ SsDs ----------------
    Public Const frmNameSsDs As String = "SsDs"
    Public Const fileNameSsDs As String = "SsDs.accdb"
    Public Const tblNameSsDs As String = "SsDs"
    Public Const fileNameUtsi As String = "UTSI.accdb"
    Public Const tblNameUtsi As String = "utsiData"
    Public Const fileNamePost As String = "post.accdb"
    Public Const tblNamePost As String = "post"
    Public Const tblNameFirm As String = "firm"
' ------------ SpecComplect ----------------
    Public Const fileNameSpecComplect As String = "specComplect.accdb"
    Public Const tblNameSpecComplData As String = "specComplData"
' ------------ Med Cod 174 ----------------
    Public Const fileNameMedCod_174 As String = "MedCod_174.accdb"
    Public Const tblNameMedCod_174 As String = "cod_174"
    Public Const tblnameCodIn_1C As String = "TDSheet"
' ------------ OKPD2 PP 1042 688 ----------------
    Public Const fileNamePpOkpd2 As String = "okpd.accdb"
    Public Const tblNamePP1042 As String = "pp1042"
    Public Const tblNamePP688 As String = "pp688"
' ------------ FTSpp312 TNVED ----------------
    Public Const fileNameFTSpp312 As String = "FTSpp312.accdb"
    Public Const tblNameFTSpp312 As String = "TNVED"
' ------------ fields Mainbase ----------------
    Public Const fldMainCod As String = "Код"
    Public Const fldMainName As String = "Наименование"

' ------------ fields Descr ----------------
    Public Const fileNameDescr As String = "Descr.accdb"
    Public Const tabNameDescr As String = "descr"
    Public Const fldDescrCod As String = "Код"
    Public Const fldDescrDesc As String = "примечание"

' ------------ fields size ----------------
    Public Const fileNameSize As String = "size.accdb"
    Public Const tblNameSize As String = "size"
    Public Const fldSizeCod As String = "Код"
    Public Const fldSizeName As String = "Наименование"
    Public Const fldSizeWidth As String = "Ширина"
    Public Const fldSizeLength As String = "Длина"
    Public Const fldSizeHeight As String = "Высота"
    Public Const fldSizeUnitLg As String = "ЕдИзмерения"
    Public Const fldSizeWeight As String = "Вес"
    Public Const fldSizeUnitWt As String = "ЕдИзмерВеса"
    Public Const fldSizeLineNumber As String = "НомерСтроки"
    Public Const fldSizeNet As String = "нетто"
    Public Const fldSizeunitNet As String = "еднНетто"
'       ---- fields in qryCountPlaceis -----
    Public Const tblCountPlaceis As String = "countPlaceis"
    Public Const fldCountPlaceisCod As String = "Код"

' ------------ KA ----------------
    Public Const fileNameKaBase As String = "KaBase.accdb"
    Public Const tabNameKaBase As String = "nomenclature"
    Public Const fldKaCod As String = "Код"
    Public Const fldKaName As String = "Наименование"

