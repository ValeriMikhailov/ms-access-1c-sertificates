VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Kazakhstan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub authorKaz_Click()
    DoCmd.OpenForm "DatesInsertKaz"
End Sub

Private Sub importKazBase_Click()
    Call StartCreateNewKazBase
End Sub

Private Sub openKazBase_Click()
    Call crtOpenKazBase
End Sub
