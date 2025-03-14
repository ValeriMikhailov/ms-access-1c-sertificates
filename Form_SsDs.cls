VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SsDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub endDsSsLeftover_Click()
    Call crtEndDsSsLeftover
End Sub
Private Sub endSs_1C_Click()
    Call crtEndSsTo1c
End Sub
Private Sub endDs_1C_Click()
    Call crtEndDsTo1c
End Sub
Private Sub endSsDs_Click()
    Call crtEndSsDs
End Sub
Private Sub endUTSI_Click()
    Call crtEndUtsi
End Sub
Private Sub fromLinks_Click()
    Call crtFromLinksToScansSertificates
End Sub
Private Sub mainDS_Click()
    Call crtMainDS
End Sub
Private Sub mainSS_Click()
    Call crtMainSS
End Sub
Private Sub mainUtsi_Click()
    Call crtMainUTSI
End Sub
Private Sub toSsDs_Click()
    Call crtToSsDs
End Sub
Private Sub toUTSI_Click()
    Call crtToUTSI
End Sub
Private Sub SsDsMovies_Click()
    Call crtSsDsMovies
End Sub
Private Sub SsDsPostDs_Click()
    Call crtSsDsPostDs
End Sub
Private Sub rplOldSsDsUt_Click()
'TODO запросы НЕ формируются кодом, при переносе проекта в другую папку путь утратится
DoCmd.RunSQL "UPDATE oldDs SET DS = 0"
DoCmd.RunSQL "UPDATE oldSs SET ID = 0"
DoCmd.RunSQL "UPDATE oldUt SET UT = 0"
End Sub
