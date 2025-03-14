VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DatesInsertKaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Result_Click()

On Error GoTo HandleError
    If IsNull(Me.fldDateBegin) Then
        Call MissingField(Me.fldDateBegin, "ÌÛÊÌ‡" & Chr(13) & "—“¿–“Œ¬¿ﬂ" & Chr(13) & "‰‡Ú‡")
        GoTo Done
    Else
        SD = DateForSQL(Me.fldDateBegin)
    End If
    
    If IsNull(Me.fldDateEnd) Then
        Call MissingField(Me.fldDateEnd, "ÌÛÊÌ‡" & Chr(13) & " ŒÕ≈◊Õ¿ﬂ" & Chr(13) & "‰‡Ú‡")
        GoTo Done
    Else
        FD = DateForSQL(Me.fldDateEnd)
    End If
    
        Call crtAuthorKazQry
Done:
    Exit Sub
HandleError:
    MsgBox "Error: " & Err.Number & vbCr & Err.Description
    Resume Done

End Sub

'---------------------------------------
Private Sub MissingField(ctl As Control, sMsg As String)
'---------------------------------------
' Purpose:  Tells user they've not filled in control <ctl>
' Requires: Reference to control; error message <sMsg>
' Returns:  --
    MsgBox sMsg, vbExclamation, "Missing Field"
    ctl.SetFocus
End Sub

Public Function DateForSQL(dt)
Dim D, M, Y As Integer
    D = Day(dt)
If Len(D) = 1 Then D = "0" & D
    M = Month(dt)
If Len(M) = 1 Then M = "0" & M
    Y = Year(dt)
DateForSQL = "#" & M & "/" & D & "/" & Y & "#"
End Function

