Attribute VB_Name = "Random"
Option Base 1
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ColorRange()
Dim nrows As Integer, ncols As Integer, blank As Single, red As Single, satisfy As Single
Dim row As Integer, col As Integer, r As Single, square As Range, response As String
Dim nblank As Integer
'MsgBox "Form is prefilled, but feel free to change", , "For the typical lazy user"

'Can also use public variable to avoid parsing values
Call nbform.allValid(nrows, ncols, blank, red, satisfy)

If blank + red >= 1 Then
    MsgBox "Sum of fraction blank and red should be smaller than 1", , "Try again"
    Call nbform.allValid(nrows, ncols, blank, red, satisfy)
End If

'unreasonable randomness
Do While satisfy = 0 Or satisfy > 0.89 Or blank = 0 Or red = 0
    Call nbform.allValid(nrows, ncols, blank, red, satisfy)
Loop
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
MsgBox "Fraction red = " & red & ", fraction blue = " & round(1 - red - blank, 2) & ", fraction blank = " & blank _
& vbNewLine & "Next comes the similarity index part", , "Fancied input"

Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)

square = fractionarr
square.Font.Size = 11

MsgBox "Now to the moving part, satisfied if cell similarity >= " & satisfy, , "Fancy fancy fancy..."
'Moving dissatisfied cells
Dim allsatisfy As Boolean
allsatisfy = True

Dim rounds As Integer, loopbound As Integer
loopbound = 100
rounds = 1
Dim dissatisfied() As String, nd As Integer, disleft As Integer, disright As Integer, blleft As Integer, blright As Integer
Dim blanks() As String, nbl As Integer, i As Integer, minlength As Integer
'Moving loop begins
Do
i = 1
nd = 1
nbl = 1
ReDim dissatisfied(nd)
ReDim blanks(nbl)
allsatisfy = False
For row = 1 To nrows
    For col = 1 To ncols
    'Find dissatisfied cell and store its position as String
    If fractionarr(row, col) < satisfy Then
        'nd = nd + 1
        ReDim Preserve dissatisfied(nd)
        dissatisfied(nd) = row & " " & col
        nd = nd + 1
            
        
    End If
    
    'Find blank cells and store positions in array
    If arrsquare(row, col) = "Blank" Then
        'nbl = nbl + 1
        ReDim Preserve blanks(nbl)
        blanks(nbl) = row & " " & col
        nbl = nbl + 1
        
        
    End If
    
    Next col
Next row
    
    'No dissatisfied neighbor. TERMINATE &&&&&&&&&&&&****************----------------------
    If nd = 1 Then
        Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)
        square = fractionarr
        MsgBox "Satisfy at round " & rounds & vbNewLine & "Similarity requirement was " & satisfy, , "Fancy Congrats"
        Exit Sub
    End If
    'avoid array out of bounds
    If nd >= nbl Then minlength = nbl Else minlength = nd
    
    'Shuffle available blank positions so the dissatisfied agent do not move to the same blank spot
    Call ShuffleArrayInPlace(blanks)
    Call ShuffleArrayInPlace(dissatisfied)
    
    Do While i < minlength
    'extract postion from string
    disleft = CInt(Split(dissatisfied(i), " ")(0))
    disright = CInt(Split(dissatisfied(i), " ")(1))

    blleft = CInt(Split(blanks(i), " ")(0))
    blright = CInt(Split(blanks(i), " ")(1))

        'Beginning to swap
        If arrsquare(disleft, disright) = "Red" Then
            'dissatisfied to blank
            arrsquare(disleft, disright) = "Blank"
            Cells(disleft, disright).Clear
            'blank to dissatisfied
            arrsquare(blleft, blright) = "Red"
            Cells(blleft, blright).Interior.color = vbRed
        ElseIf arrsquare(disleft, disright) = "Blue" Then
            'dissatisfied to blank
            arrsquare(disleft, disright) = "Blank"
            Cells(disleft, disright).Clear
            'blank to dissatisfied
            arrsquare(blleft, blright) = "Blue"
            Cells(blleft, blright).Interior.color = vbBlue
            
        End If
        
        i = i + 1
        
    Loop
    'Application.Wait (Now + TimeValue("0:00:1"))

    Application.ScreenUpdating = True
    


'Recalculate similarity for next round
Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)
'square = fractionarr

rounds = rounds + 1

If rounds > loopbound Then
    Call SimilarityIndex(nrows, ncols, arrsquare, fractionarr)
    response = MsgBox("More than " & loopbound & " rounds have passed. It's possible that not every neighbor can be satisfied" _
    & vbNewLine & "Similarity threshold was " & satisfy & vbNewLine & "Do you wish to continue?", vbYesNo, "Continue?")
    If response = vbYes Then loopbound = loopbound + loopbound
    If response = vbNo Then
        square = fractionarr
        Exit Sub
    End If


End If

Loop Until allsatisfy



MsgBox "All happy at round " & rounds & vbNewLine & "Similarity threshold was " & satisfy, , "Fancy congrats"

End Sub

Sub SimilarityIndex(nrows As Integer, ncols As Integer, myArray() As String, fractionArray() As Double)
    Dim i As Integer, J As Integer, likecount As Integer, notlike As Integer
    Dim row As Integer, Column As Integer

    ReDim fractionArray(1 To nrows, 1 To ncols) As Double

    For i = 1 To nrows
        For J = 1 To ncols
            fractionArray(i, J) = 0
        Next
    Next

    For row = 1 To nrows
        For Column = 1 To ncols
        likecount = 0
        notlike = 0

        If myArray(row, Column) = "Blank" Then
            fractionArray(row, Column) = 99
            'Cells(row, Column).Value = 99
        Else
            For i = -1 To 1
                For J = -1 To 1
                If row + i <= 0 Or Column + J <= 0 Or row + i > nrows Or Column + J > ncols Then
                ElseIf myArray(row + i, Column + J) = "Red" Or myArray(row + i, Column + J) = "Blue" Then
                    If i = 0 And J = 0 Then
                    ElseIf myArray(row + i, Column + J) = myArray(row, Column) Then
                        likecount = likecount + 1
                    Else
                        notlike = notlike + 1
                    End If
                End If
                Next J
            Next i
            If (likecount + notlike) = 0 Then
                fractionArray(row, Column) = 1 'all blank neighbors: 0 if dissatisfied, 1 if satisfied -> your choice
            Else
                fractionArray(row, Column) = (likecount / (likecount + notlike))
            End If
        End If
        Next Column
    Next row

'    For row = 1 To nrows
'        For Column = 1 To ncols
'            If Not fractionArray(row, Column) = 99 Then
'                Cells(row, Column) = round(fractionArray(row, Column), 2)
'                Cells(row, Column).Font.ColorIndex = 6
'            End If
'        Next Column
'    Next row

End Sub
'--------------------------------UNUSED----------------------------------------------------
Function cellSimilar(myArray() As String, nrows As Integer, ncols As Integer, row As Integer, col As Integer) As Double
Dim i As Integer, J As Integer, likecount As Integer, notlike As Integer
    For i = -1 To 1
        For J = -1 To 1
        If row + i <= 0 Or col + J <= 0 Or row + i > nrows Or col + J > ncols Then
        ElseIf myArray(row + i, col + J) = "Red" Or myArray(row + i, col + J) = "Blue" Then
            If i = 0 And J = 0 Then
            ElseIf myArray(row + i, col + J) = myArray(row, col) Then
                likecount = likecount + 1
            Else
                notlike = notlike + 1
            End If
        End If
        Next J
    Next i
    If (likecount + notlike) = 0 Then
        cellSimilar = 1
        'fractionArray(row, Column) = 1 'all blank neighbors: 0 if dissatisfied, 1 if satisfied -> your choice
    Else
        cellSimilar = (likecount / (likecount + notlike))
        'fractionArray(row, Column) = (likecount / (likecount + notlike))
    End If

End Function


Sub ShuffleArrayInPlace(InArray() As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArrayInPlace
' This shuffles InArray to random order, randomized in place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
   
    Randomize
    For N = LBound(InArray) To UBound(InArray)
        J = CLng(((UBound(InArray) - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
End Sub
