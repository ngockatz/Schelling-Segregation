'Macro for Model 2
Option Base 1
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ColorRange()
Dim nrows As Integer, ncols As Integer, blank As Single, red As Single, satisfy As Single
Dim row As Integer, col As Integer, r As Single, square As Range, response As String
Dim nblank As Integer
Dim sleeptime As Integer

With nbform
    Call .Initialize
    .Show
    nrows = .sbnRows.Value
    ncols = .sbnCols.Value
    blank = .sbBlank.Value / 100
    red = .sbRed.Value / 100
    satisfy = .sbSatisfy.Value / 100
    sleeptime = .sbDelay.Value
End With

ReDim arrsquare(nrows, ncols) As String
ReDim fractionarr(nrows, ncols) As Double

Set square = Range(Cells(1, 1), Cells(nrows, ncols))

ActiveSheet.Cells.Clear

'Divide probability into range of x, x+red, x + red + blue
For row = 1 To nrows
    For col = 1 To ncols

    r = Rnd()
    If r < blank Then
        arrsquare(row, col) = "Blank"
        nblank = nblank + 1
    ElseIf r >= blank And r < red + blank Then
        arrsquare(row, col) = "Red"
        square(row, col).Interior.color = vbRed
    Else
        arrsquare(row, col) = "Blue"
        square(row, col).Interior.color = vbBlue
    End If
    Next col
Next row

square = arrsquare

MsgBox "Approximately = " & red * 100 & "% is red, " & round(1 - red - blank, 2) * 100 & "% is blue, " & _
blank * 100 & "% is blank" & vbNewLine _
& "Next take a look at the current satisfaction rate", , "Neighbor input"

Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)

square = fractionarr
square.Font.Size = 11

MsgBox "Look at the calculated grid. Note that blank is arbitrarily set to 99" & vbNewLine _
& "Now let's move dissatisfied agents to any vacant spot, keep moving until the neighbor similarity is at least " & satisfy * 100 & "%", , _
"Moving dissatisfied folks"

'Moving dissatisfied cells
Dim allsatisfy As Boolean, blankchecked As Integer
allsatisfy = True

Dim y As Integer, z As Integer, step As Integer, rounds As Integer, a As Integer, b As Integer, loopbound As Integer
Dim blanknextround() As Integer
loopbound = 25
rounds = 0

Do
allsatisfy = True
For row = 1 To nrows
    For col = 1 To ncols

    'Find a dissatisfied cell
    If Not fractionarr(row, col) >= satisfy Then
        allsatisfy = False
        step = 0
        'Find a blank cell
        'Start from beginning
        For a = 1 To nrows
        For b = 1 To ncols

            If arrsquare(a, b) = "Blank" Then
                'Swap label
            
                arrsquare(a, b) = arrsquare(row, col)

                arrsquare(row, col) = "Blank"

                'check if satisfied at new place
                If cellSimilar(arrsquare, nrows, ncols, a, b) < satisfy Then
 
                    'reverse as before
                    arrsquare(row, col) = arrsquare(a, b)
                    arrsquare(a, b) = "Blank"

                Else
                    Sleep sleeptime
                    arrsquare(row, col) = "Blank"
                    fractionarr(a, b) = cellSimilar(arrsquare, nrows, ncols, a, b)
                    Cells(a, b).Value = fractionarr(a, b)
                    If arrsquare(a, b) = "Red" Then Cells(a, b).Interior.color = vbRed
                    If arrsquare(a, b) = "Blue" Then Cells(a, b).Interior.color = vbBlue
                    
                    'unsatisfied cells now become blank
                    fractionarr(row, col) = 99
                    Cells(row, col).Clear

                    'found 1 feasible location
                    step = 1
                End If
                
                If step > 0 Then Exit For
            End If
        Next b
        'exit outer loop early
        If step > 0 Then Exit For

        Next a
        
    End If

    Next col
Next row


'Recalculate similarity for next round
Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)

Application.ScreenUpdating = True

rounds = rounds + 1
If rounds > loopbound Then
    response = MsgBox("More than " & loopbound & " rounds have passed. It's possible that not every neighbor can be satisfied" _
    & vbNewLine & "Similarity threshold was " & satisfy & vbNewLine & "Do you wish to continue?", vbYesNo, "Continue?")
    If response = vbYes Then loopbound = loopbound + loopbound
    If response = vbNo Then Exit Sub
End If

Loop Until allsatisfy

MsgBox "All agents satisfied at round " & rounds & vbNewLine _
        & "Similarity requirement was " & satisfy * 100 & "%" & vbNewLine _
        & "Feel free to inspect the calculated grid", , "Summary"

square = fractionarr

End Sub

Sub SimilarityIndex(nrows As Integer, ncols As Integer, myArray() As String, fractionArray() As Double)
    Dim i As Integer, j As Integer, likecount As Integer, notlike As Integer
    Dim row As Integer, Column As Integer

    ReDim fractionArray(1 To nrows, 1 To ncols) As Double

    For i = 1 To nrows
        For j = 1 To ncols
            fractionArray(i, j) = 0
        Next
    Next

    For row = 1 To nrows
        For Column = 1 To ncols
        likecount = 0
        notlike = 0

        If myArray(row, Column) = "Blank" Then
            fractionArray(row, Column) = 99
        Else
            For i = -1 To 1
                For j = -1 To 1
                If row + i <= 0 Or Column + j <= 0 Or row + i > nrows Or Column + j > ncols Then
                ElseIf myArray(row + i, Column + j) = "Red" Or myArray(row + i, Column + j) = "Blue" Then
                    If i = 0 And j = 0 Then
                    ElseIf myArray(row + i, Column + j) = myArray(row, Column) Then
                        likecount = likecount + 1
                    Else
                        notlike = notlike + 1
                    End If
                End If
                Next j
            Next i
            If (likecount + notlike) = 0 Then
                fractionArray(row, Column) = 1 'all blank neighbors: 0 if dissatisfied, 1 if satisfied -> your choice
            Else
                fractionArray(row, Column) = (likecount / (likecount + notlike))
            End If
        End If
        Next Column
    Next row

    For row = 1 To nrows
        For Column = 1 To ncols
            If Not fractionArray(row, Column) = 99 Then
                Cells(row, Column) = round(fractionArray(row, Column), 2)
                Cells(row, Column).Font.ColorIndex = 6
            End If
        Next Column
    Next row

End Sub

Function cellSimilar(myArray() As String, nrows As Integer, ncols As Integer, row As Integer, col As Integer) As Double
Dim i As Integer, j As Integer, likecount As Integer, notlike As Integer
    For i = -1 To 1
        For j = -1 To 1
        If row + i <= 0 Or col + j <= 0 Or row + i > nrows Or col + j > ncols Then
        ElseIf myArray(row + i, col + j) = "Red" Or myArray(row + i, col + j) = "Blue" Then
            If i = 0 And j = 0 Then
            ElseIf myArray(row + i, col + j) = myArray(row, col) Then
                likecount = likecount + 1
            Else
                notlike = notlike + 1
            End If
        End If
        Next j
    Next i
    If (likecount + notlike) = 0 Then
        cellSimilar = 1
    Else
        cellSimilar = (likecount / (likecount + notlike))
    End If

End Function

