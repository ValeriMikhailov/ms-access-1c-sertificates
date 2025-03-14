VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub articul_Click()
    Call crtArticul
End Sub
Private Sub articulSplit_Click()
    Call crtArticulSplit
End Sub
Private Sub author_Click()
    DoCmd.OpenForm frmNameDatesInsert
'    Call crtAuthorQry
End Sub
Private Sub exportCSV_Click()
'из ACCESS выгружает с кавычками, приходится выгружать через EXCEL
    Dim excelApp
    Call crtCsvRu
    Call crtCsvDs
    Call crtCsvSs
    Call crtCsvUtsi
    
    Call ExportTableInToExcelFile("csvRU", pathMainBase & fileNameExcelToCSV)
    Call ExportTableInToExcelFile("csvDS", pathMainBase & fileNameExcelToCSV)
    Call ExportTableInToExcelFile("csvSS", pathMainBase & fileNameExcelToCSV)
    Call ExportTableInToExcelFile("csvUtsi", pathMainBase & fileNameExcelToCSV)
    
    nameExcelMacro = nameExcelMacroExportToCSV
    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   'макрос Excel exportToCSV
End Sub
Private Sub FTSpp312_Click()
    Call crtFTSpp312
End Sub

Private Sub getUTSI_Click()
    Call crtGetUtsi
End Sub
Private Sub GTD_Click()
    Call crtGTDbase
End Sub

Private Sub importGTD_Click()
    Call StartCreateNewGTDbase
End Sub

Private Sub importReceipts_Click()
    Call StartCreateNewReceipts
End Sub
Private Sub importSize_Click()
    Call StartCreateNewSize
End Sub

Private Sub Kazakhstan_Click()
    DoCmd.OpenForm "Kazakhstan"
End Sub

Private Sub LinksToSCANS_Click()
    If SearchOK() Then
        Call LinksToScansRegUd
    Else
        Call LinksToScanSertificates
    End If
End Sub
Private Sub LinksToScansResult_Click()
    Call LinksUpdateToScansResult
End Sub
Private Sub mainDbEndSsDs_Click()
    Call crtMainDbEndSsDs
End Sub
Private Sub MED_Cod_174_Click()
    Call crtMedCod
End Sub
Private Sub MED_Cod_1C_Click()
    Call crtMedCod_1c
End Sub
Private Sub otkaz_Click()
    Call crtSplitOtkaz
    DoCmd.OpenForm frmnameOtkaz
End Sub
Private Sub Post_Click()
    Call crtPost
End Sub
Private Sub getPP1042_Click()
    Call crtPP1042
End Sub
Private Sub pp688_Click()
    Call crtPP688
End Sub
Private Sub receiptsRegUd_Click()
'    Call crtReceiptsRegUd
    DoCmd.OpenForm frmnameReceiptsRegUd
End Sub
Private Sub sizeBrutto_Click()
    Call crtSizeBrutto
End Sub
Private Sub sizeDescr_Click()
    Call crtSizeDescr
End Sub
Private Sub specComplData_Click()
    Call crtSpecComplData
End Sub
Private Sub specComplResult_Click()
    Call crtSpecComplResult
End Sub
Private Sub SplitRow_Click()
    Call crtSplitRow
End Sub
Private Sub SsDs_Click()
    DoCmd.OpenForm frmNameSsDs
End Sub
Private Sub templateQuery_Click()
    Call crtTemplateQry
End Sub
Private Sub StartCreateNewDatabaseAndImportExcelTable_Click()
    Call StartCreateNewDBandImportExcelTable
End Sub
Private Sub StartCreateNewDatabaseAndImportRegUd_Click()
    Call ImportRegUd
End Sub
Private Sub descr_Click()
    Call crtDescr
End Sub
Private Sub doubleRows_Click()
    Call crtDoubleRows
End Sub
Private Sub Folders_Click()
    Call crtFolders
    DoCmd.OpenForm frmnameFolders
End Sub
Private Sub getRu_Click()
    Call crtGetRu
End Sub
Private Sub getDS_Click()
    Call crtGetDs
End Sub
Private Sub getSS_Click()
    Call crtGetSs
End Sub
Private Sub main0_Click()
    Call crtMain
End Sub
Private Sub main1_Click()
    Call crtMain1
End Sub
Private Sub main2_Click()
    Call crtMain
End Sub
Private Sub MaunBase_Click()
    Call crtOpenMainNomenclature
End Sub
Private Sub endRegUd_Click()
    Call crtDateEndRegUd
End Sub
Private Sub pathRegUd_Click()
    Call crtPathRegUd
    DoCmd.OpenForm frmnamePathRegUd
End Sub
Private Sub pathSertificates_Click()
    Call crtPathSertificates
    DoCmd.OpenForm frmnamePathSertificates
End Sub
Private Sub produce_Click()
    Call crtProduce
End Sub
Private Sub regUdDescr_Click()
    Call crtRegUdDescr
End Sub
Private Sub regUdSearch_0_Click()
    Call crtRegUdSearch
End Sub
Private Sub regUdSearch1_Click()
    Call crtRegUdSearch1
End Sub
Private Sub regUdSearch2_Click()
    Call crtRegUdSearch2
End Sub
Private Sub seller_Click()
    Call crtSeller
    DoCmd.OpenForm frmnameSeller
End Sub
Private Sub showEditExportTo1C_Click()
    Call crtExportTo1C
End Sub
Private Sub specComplectTextDescr_Click()
    Call crtSpecComplectQry
End Sub
Private Sub tnved_Click()
'    LoadFromText acModule, "fncTNVED", "D:\accModulesExportImport\fncTNVED.bas"
'    Modules("fncTNVED").AddFromFile ("D:\accModulesExportImport\fncTNVED.bas")
    Call crtTNVED
End Sub
Private Sub tnved0_Click()
    Call crtTNVED1
End Sub
Private Sub tnvedALL_Click()
    Call crtTnvedALL
End Sub
Private Sub changesRegUd_Click()
    Call crtChangesRegUd
End Sub
Private Sub okpRegUd_Click()
    Call crtOkpRegUd
'    DoCmd.OpenQuery "okpRegUd", , acEdit
End Sub
Private Sub exportToServerExcel_Click()
    Dim excelApp
    Call ExportTableInToExcelFile("getRu", pathToServer & fileNameExcelOnServer)
    Call ExportTableInToExcelFile("pathRegUd", pathToServer & fileNameExcelOnServer)
    Call ExportTableInToExcelFile("getDS", pathToServer & fileNameExcelOnServer)
    Call ExportTableInToExcelFile("getSS", pathToServer & fileNameExcelOnServer)
    Call ExportTableInToExcelFile("pathSertificates", pathToServer & fileNameExcelOnServer)
    
    nameExcelMacro = nameExcelMacroFormatingCells
    excelApp = StartExcelFileMacro(pathMainBase & fileExcelMacro, nameExcelMacro)   'макрос Excel changesForEnd
End Sub
