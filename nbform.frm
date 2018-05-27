*** Form front end properties ***
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nbform 
   Caption         =   "Neighborhood Inputs"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7920
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

Private Sub Initialize()

    tbnRows = WorksheetFunction.RandBetween(5, 15)
    tbnCols = WorksheetFunction.RandBetween(5, 15)
    tbRed = round(Rnd(), 1)
    If tbRed > 0.5 Then tbRed = 1 - tbRed + 0.1
    tbBlank = round(Rnd(), 1)
    If tbBlank > 0.5 Then tbBlank = 1 - tbBlank + 0.1
    tbSatisfy = round(Rnd(), 1)

End Sub

Private Sub cmdOK_Click()
If Valid Then Me.Hide
End Sub
'Function to filter user input for each field
Private Function Valid() As Boolean

    Valid = True
    
    If tbnRows.Value = "" Or Not IsNumeric(tbnRows.Value) Or tbnRows.Value < 1 Then
        Valid = False
        MsgBox "Number of rows should be a whole number greater than 0", , "Oh come on"
        tbnRows.SetFocus
    End If
    
    If tbnCols.Value = "" Or Not IsNumeric(tbnCols.Value) Or tbnCols.Value < 1 Then
        Valid = False
        MsgBox "Number of columns should be a whole number greater than 0", , "Oh come on"
        tbnCols.SetFocus
    End If
    
    If tbBlank.Value = "" Or Not IsNumeric(tbBlank.Value) Or tbBlank.Value > 1 Or tbBlank.Value < 0 Then
        Valid = False
        MsgBox "Fraction of blank cells should be a decimal value between 0 and 1", , "Oh come on"
        tbBlank.SetFocus
    End If
    
    If tbRed.Value = "" Or Not IsNumeric(tbRed.Value) Or tbRed.Value > 1 Or tbRed.Value < 0 Then
        Valid = False
        MsgBox "Fraction of red cells should be a decimal between 0 and 1", , "Oh come on"
        tbRed.SetFocus
    End If
    
    If tbSatisfy.Value = "" Or Not IsNumeric(tbSatisfy.Value) Or tbSatisfy.Value > 1 Or tbSatisfy.Value < 0 Then
        Valid = False
        MsgBox "Satisfaction rate should be a decimal between 0 and 1", , "Oh come on"
        tbSatisfy.SetFocus
    End If

End Function

Public Function allValid(nrows As Integer, ncols As Integer, blank As Single, _
red As Single, satisfy As Single) As Boolean
    
    'Initialize form's fields' as first appearing
    Call Initialize
    Me.Show
    
    'Store inputs to variables which later can be called in the main program
    nrows = tbnRows.Text
    ncols = tbnCols.Text
    blank = tbBlank.Text
    red = tbRed.Text
    satisfy = tbSatisfy.Text
    
End Function

Private Sub tbnRows_Change()

End Sub

'Generic code if user clicks X
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub
