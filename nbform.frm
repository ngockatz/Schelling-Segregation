*** Form frontend properties ***
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nbform 
   Caption         =   "Neighborhood Inputs"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8880
   OleObjectBlob   =   "nbform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "nbform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

*** Form backend ***
Private Sub cmdCancel_Click()

    Unload Me
    End

End Sub

Public Sub Initialize()

    lbRow = sbnRows.Value
    lbCol = sbnCols.Value
    lbBlank = sbBlank.Value & "%"
    sbRed.Max = 100 - sbBlank.Value
    lbRB = sbRed.Value & "/" & sbRed.Max - sbRed.Value & " %"
    lbSatisfy = sbSatisfy.Value & "%"
    lbDelay = "Animation delay: " & sbDelay.Value & " milliseconds"

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub



Private Sub sbBlank_Change()
lbBlank = sbBlank.Value & "%"
sbRed.Max = 100 - sbBlank.Value
lbRB = sbRed.Value & "/" & (sbRed.Max - sbRed.Value) & " %"
End Sub

Private Sub sbDelay_Change()
lbDelay = "Animation delay: " & sbDelay.Value & " milliseconds"
End Sub

Private Sub sbnCols_Change()
lbCol = sbnCols.Value
End Sub

Private Sub sbnRows_Change()
lbRow = sbnRows.Value
End Sub

Private Sub sbRed_Change()
lbRB = sbRed.Value & "/" & (sbRed.Max - sbRed.Value) & " %"
End Sub


Private Sub sbSatisfy_Change()
lbSatisfy = sbSatisfy.Value & "%"
End Sub


'if user clicks X
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub
