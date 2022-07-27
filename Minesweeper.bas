Attribute VB_Name = "Minesweeper"
Option Explicit

' Establishes a 10 x 10 array that corresponds to the playing field with a one cell buffer all around to prevent out-of-range errors.
' Zeros will denote no bomb; ones will denote bomb.
Dim Minefield(9, 9) As Integer
' Establishes an 8 x 8 array that corresponds to the playing field.
' Numbers will indicate the number of bombs in surrounding cells.
Dim BombCountField(7, 7) As Integer
' Establishes a 10 x 10 array that corresponds to the playing field with a one cell buffer all around to prevent out-of-range errors.
' This will be populated with ones that turn into zeros as the player reveals hidden cells.
' When the sum of this array matches the sum of the Minefield array and the remaining bomb count is zero, the player wins.
Dim WinField(9, 9) As Integer
' This will be used to indicate whether a player's move is the first move.
Dim SubsequentMove As Boolean

Sub Reset()
Attribute Reset.VB_ProcData.VB_Invoke_Func = "v\n14"

'This sub resets the playing field after a win or loss, or when the player decides to reset.

ActiveSheet.Unprotect

'Resets the unfound bomb count to 10.
Range("D1") = 10

'Resets the playing field to gray interior color, black font with Calibri style, and clears the numbers and images.
With Range("C3:J10")
    .Interior.Color = RGB(200, 200, 200)
    .Font.Color = vbBlack
    .Font.Name = "Calibri"
    .ClearContents
End With

' Resets the face in J1 to a happy face.
With Range("J1")
    .Font.Name = "Wingdings"
    .Value = "J"
End With

' Ensures that the very first move (poke) after reset will follow the first move protocol of calling Seed, rather than Poke.
SubsequentMove = False

ActiveSheet.Protect

End Sub

Sub Play()
Attribute Play.VB_ProcData.VB_Invoke_Func = "z\n14"

' Ensures that the player only selects one cell, and that the selected cell is within the playing field.
' Also determines whether to call Seed or Poke, depending on whether this move is the first move.


Dim RowInRange As Boolean
Dim ColInRange As Boolean
Dim CellInRange As Boolean


ActiveSheet.Unprotect

' Ensures that player has selected only one cell.
If Selection.Cells.Count = 1 Then

    ' Ensures that selected cell is within the playing field.
    If ActiveCell.Row >= 3 And ActiveCell.Row <= 10 Then
        RowInRange = True
    Else
        RowInRange = False
    End If
    
    If ActiveCell.Column >= 3 And ActiveCell.Column <= 10 Then
        ColInRange = True
    Else
        ColInRange = False
    End If
    
    If RowInRange And ColInRange Then
        If SubsequentMove = False Then
            Call Seed
        ElseIf SubsequentMove = True Then
            Call Poke
        End If
    Else
        MsgBox ("Select a cell within the minefield.")
    End If
    
Else
    MsgBox ("Select only one cell within the minefield.")
End If

ActiveSheet.Protect

End Sub

Sub Seed()

' This sub establishes a Minefield array that guarantees a safe first move,
' and a WinField array that will have to match the Minefield array to win the game.


Dim Col As Integer
Dim Row As Integer
Dim RandCount As Integer
Dim RowRand As Integer
Dim ColRand As Integer
Dim SelectRow As Boolean
Dim SelectCol As Boolean
Dim SelectCell As Boolean
Dim BombCount As Integer
Dim NeighRow As Integer
Dim NeighCol As Integer
Dim Self As Boolean


'Establishes empty Minefield array.
For Col = 0 To 9
    For Row = 0 To 9
        Minefield(Row, Col) = 0
    Next Row
Next Col

'Establishes initial WinField array.
For Col = 0 To 9
    For Row = 0 To 9
        WinField(Row, Col) = 0
    Next Row
Next Col

For Col = 1 To 8
    For Row = 1 To 8
        WinField(Row, Col) = 1
    Next Row
Next Col

'Ensures that no bomb is placed on or adjacent to the first selected cell.
SelectCell = False

'Places 10 bombs randomly in Minefield array while avoiding the slected cell and cells adjacent to the selected cell.
RandCount = 0

Do Until RandCount = 10
    RowRand = Int(1 + Rnd * (8 - 1 + 1))
    ColRand = Int(1 + Rnd * (8 - 1 + 1))
    
    If RowRand + 2 >= ActiveCell.Row - 1 And RowRand + 2 <= ActiveCell.Row + 1 Then
        SelectRow = True
    Else
        SelectRow = False
    End If

    If ColRand + 2 >= ActiveCell.Column - 1 And ColRand + 2 <= ActiveCell.Column + 1 Then
        SelectCol = True
    Else
        SelectCol = False
    End If
    
    If SelectRow And SelectCol Then
        SelectCell = True
    Else
        SelectCell = False
    End If
    
    If Minefield(RowRand, ColRand) = 0 And SelectCell = False Then
        Minefield(RowRand, ColRand) = 1
        RandCount = RandCount + 1
    End If
Loop

'For cheating or code testing, displays the Minefield array.

'For Col = 1 To 8
'    For Row = 1 To 8
'        Cells(Row + 2, Col + 12) = Minefield(Row, Col)
'    Next Row
'Next Col

'Populates BombCountField array, which contains the number of bombs surrounding a cell.
For Col = 1 To 8
    For Row = 1 To 8
        BombCount = 0
        
        For NeighRow = -1 To 1
            For NeighCol = -1 To 1
            
                If NeighRow = 0 And NeighCol = 0 Then
                    Self = True
                Else
                    Self = False
                End If
                
                If Self = False Then
                    BombCount = BombCount + Minefield(Row + NeighRow, Col + NeighCol)
                End If
                
            Next NeighCol
        Next NeighRow
        
        BombCountField(Row - 1, Col - 1) = BombCount
    Next Row
Next Col

'For cheating or code testing, displays the BombCountField array.

'For Col = 1 To 8
'    For Row = 1 To 8
'        Cells(Row + 12, Col + 12) = BombCountField(Row - 1, Col - 1)
'    Next Row
'Next Col

'Marks all future moves as not the first move.
SubsequentMove = True

Call Poke

End Sub

Sub Poke()

' Determines whether the selected cell has a bomb and responds appropriately.
' When a cell with no bomb is poked, the number of surrounding bombs for that cell will be revealed.
' If the poked cell has no adjacent bombs, adjacent cells will also be cleared until reaching cells with adjacent bombs.


Dim NeighCol As Integer
Dim NeighRow As Integer
Dim RowInRange As Boolean
Dim ColInRange As Boolean
Dim CellInRange As Boolean
Dim Square As Range


'If cell is already flagged, Poke will not execute.
If ActiveCell.Value <> "P" Then
    'Calls a routine to temporarily display a serious face in J1 long enough to be perceived by player.
    Call SeriousFace(1)
    'Losing sequence for poking a cell with a bomb.
    If Minefield(ActiveCell.Row - 2, ActiveCell.Column - 2) = 1 Then
        'Changes the face in J1 to an unhappy face.
        With Range("J1")
            .Font.Name = "Wingdings"
            .Value = "L"
        End With
        For Each Square In Range("C3:J10")
            'Reveals hidden, unflagged bombs.
            If Minefield(Square.Row - 2, Square.Column - 2) = 1 And Square.Value <> "P" Then
                Square.Interior.Color = vbWhite
                Square.Font.Name = "Wingdings"
                Square.Font.ColorIndex = 9
                Square.Value = "M"
            Else
                'Replaces incorrectly placed flags with "X"s.
                If Minefield(Square.Row - 2, Square.Column - 2) = 0 And Square.Value = "P" Then
                    Square.Font.Name = "Wingdings 2"
                    Square.Font.Color = vbBlack
                    Square.Value = "Ñ"
                End If
            End If
        Next Square
        ActiveCell.Interior.Color = vbRed
        MsgBox (":(  Game Over.")
        Call Reset
    'Sequence for poking cell without a bomb within it or adjacent to it.  Will reveal that cell and adjacent cells.
    ElseIf BombCountField(ActiveCell.Row - 3, ActiveCell.Column - 3) = 0 Then
        For NeighRow = -1 To 1
            For NeighCol = -1 To 1
            
                'Ensures no out-of-range errors.
                If ActiveCell.Row + NeighRow >= 3 And ActiveCell.Row + NeighRow <= 10 Then
                    RowInRange = True
                Else
                    RowInRange = False
                End If
                
                If ActiveCell.Column + NeighCol >= 3 And ActiveCell.Column + NeighCol <= 10 Then
                    ColInRange = True
                Else
                    ColInRange = False
                End If
                
                If RowInRange And ColInRange Then
                    CellInRange = True
                Else
                    CellInRange = False
                End If
            
                If CellInRange And Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Value <> "P" Then
            
                    'Turns revealed and unflagged cells white.
                    Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Interior.Color = vbWhite
                    'Changes corresponding WinField number to 0.
                    WinField(ActiveCell.Row + NeighRow - 2, ActiveCell.Column + NeighCol - 2) = 0
                    'Ensures font is not Winddings.
                    Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.Name = "Calibri"
                    
                    'Bases font color on adjacent bomb count.
                    Select Case BombCountField(ActiveCell.Row + NeighRow - 3, ActiveCell.Column + NeighCol - 3)
                        Case 1
                            Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.Color = vbBlue
                        Case 2
                            Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.Color = vbGreen
                        Case 3
                            Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.Color = vbRed
                        Case 4
                            Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.ColorIndex = 11
                        Case 5
                            Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol).Font.ColorIndex = 9
                    End Select
                    
                    'Displays non-zero adjacent bomb count.
                    If BombCountField(ActiveCell.Row + NeighRow - 3, ActiveCell.Column + NeighCol - 3) <> 0 Then
                        Cells(ActiveCell.Row + NeighRow, ActiveCell.Column + NeighCol) = BombCountField(ActiveCell.Row + NeighRow - 3, ActiveCell.Column + NeighCol - 3)
                    End If
            
                End If
            
            Next NeighCol
        Next NeighRow
        
        'Resets the face in J1 to a happy face.
        With Range("J1")
            .Font.Name = "Wingdings"
            .Value = "J"
        End With
            
        'Winning sequence for revealing all cells without bombs and correctly flagging all cells with bombs.
        If WorksheetFunction.Sum(WinField) = WorksheetFunction.Sum(Minefield) And Range("D1").Value = 0 Then
            MsgBox ("Congratulations!!! You win!!!")
            Call Reset
        End If
        
        'Calls routine to further clear areas around freshly revealed cells with zero adjacent bombs.
        Call Repoke
        
    'Sequence for poking cell without a bomb within it, but with at least one bomb adjacent to it.  Will reveal only that cell.
    Else
        'Turns revealed and unflagged cells white.
        ActiveCell.Interior.Color = vbWhite
        'Changes corresponding WinField number to 0.
        WinField(ActiveCell.Row - 2, ActiveCell.Column - 2) = 0
        'Ensures font is not Winddings.
        ActiveCell.Font.Name = "Calibri"
        
        'Bases font color on adjacent bomb count.
        Select Case BombCountField(ActiveCell.Row - 3, ActiveCell.Column - 3)
            Case 1
                ActiveCell.Font.Color = vbBlue
            Case 2
                ActiveCell.Font.Color = vbGreen
            Case 3
                ActiveCell.Font.Color = vbRed
            Case 4
                ActiveCell.Font.ColorIndex = 11
            Case 5
                ActiveCell.Font.ColorIndex = 9
        End Select
        
        'Displays non-zero adjacent bomb count.
        ActiveCell = BombCountField(ActiveCell.Row - 3, ActiveCell.Column - 3)
        
        'Resets the face in J1 to a happy face.
        With Range("J1")
            .Font.Name = "Wingdings"
            .Value = "J"
        End With
            
        'Winning sequence for revealing all cells without bombs and correctly flagging all cells with bombs.
        If WorksheetFunction.Sum(WinField) = WorksheetFunction.Sum(Minefield) And Range("D1").Value = 0 Then
            SubsequentMove = False
            MsgBox ("Congratulations!!! You win!!!")
            Call Reset
        End If
    
    End If
    
End If

End Sub

Sub Repoke()

' Ensures that freshly revealed cells with zero adjacent bombs also get their surrounding cells
' revealed in order to make gameplay less manual.


Dim Square As Range
Dim NeighCol As Integer
Dim NeighRow As Integer
Dim RowInRange As Boolean
Dim ColInRange As Boolean
Dim CellInRange As Boolean
Dim ZeroCountOld As Integer
Dim ZeroCountNew As Integer
Dim FoundAll As Boolean


ZeroCountOld = 0
ZeroCountNew = 0

Do While FoundAll = False

    For Each Square In Range("C3:J10")
    
        If Square.Interior.Color = vbWhite And BombCountField(Square.Row - 3, Square.Column - 3) = 0 Then
        
        ZeroCountNew = ZeroCountNew + 1
        
            For NeighRow = -1 To 1
                For NeighCol = -1 To 1
                
                    'Ensures no out-of-range errors.
                    If Square.Row + NeighRow >= 3 And Square.Row + NeighRow <= 10 Then
                        RowInRange = True
                    Else
                        RowInRange = False
                    End If
                    
                    If Square.Column + NeighCol >= 3 And Square.Column + NeighCol <= 10 Then
                        ColInRange = True
                    Else
                        ColInRange = False
                    End If
                    
                    If RowInRange And ColInRange Then
                        CellInRange = True
                    Else
                        CellInRange = False
                    End If
                
                    If CellInRange And Cells(Square.Row + NeighRow, Square.Column + NeighCol).Value <> "P" Then
                
                        Cells(Square.Row + NeighRow, Square.Column + NeighCol).Interior.Color = vbWhite
                        WinField(Square.Row + NeighRow - 2, Square.Column + NeighCol - 2) = 0
                        Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.Name = "Calibri"
                        
                        Select Case BombCountField(Square.Row + NeighRow - 3, Square.Column + NeighCol - 3)
                            Case 1
                                Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.Color = vbBlue
                            Case 2
                                Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.Color = vbGreen
                            Case 3
                                Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.Color = vbRed
                            Case 4
                                Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.ColorIndex = 11
                            Case 5
                                Cells(Square.Row + NeighRow, Square.Column + NeighCol).Font.ColorIndex = 9
                        End Select
                        
                        If BombCountField(Square.Row + NeighRow - 3, Square.Column + NeighCol - 3) <> 0 Then
                            Cells(Square.Row + NeighRow, Square.Column + NeighCol) = BombCountField(Square.Row + NeighRow - 3, Square.Column + NeighCol - 3)
                        End If
                
                    End If
                
                Next NeighCol
            Next NeighRow
                
            'Winning sequence for revealing all cells without bombs and correctly flagging all cells with bombs.
            If WorksheetFunction.Sum(WinField) = WorksheetFunction.Sum(Minefield) And Range("D1").Value = 0 Then
                MsgBox ("Congratulations!!! You win!!!")
                Call Reset
            End If
        
        End If
    
    Next Square
    
    If ZeroCountOld = ZeroCountNew Then
        FoundAll = True
    Else
        FoundAll = False
    End If
    
    ZeroCountOld = ZeroCountNew
    
    ZeroCountNew = 0
    
Loop

End Sub

Sub Flag()
Attribute Flag.VB_ProcData.VB_Invoke_Func = "x\n14"

' Allows the player to flag suspected cells with bombs and safeguard them from poking.


Dim RowInRange As Boolean
Dim ColInRange As Boolean
Dim CellInRange As Boolean


ActiveSheet.Unprotect

'Ensures player can only flag cells in the playing field.
If ActiveCell.Row >= 3 And ActiveCell.Row <= 10 Then
    RowInRange = True
Else
    RowInRange = False
End If

If ActiveCell.Column >= 3 And ActiveCell.Column <= 10 Then
    ColInRange = True
Else
    ColInRange = False
End If

If RowInRange And ColInRange Then
    'Ensures player can only flag unrevealed cells.
    If ActiveCell.Interior.Color = RGB(200, 200, 200) Then
        'Adds a flag if no flag exists.
        If ActiveCell <> "P" Then
            ActiveCell.Font.Color = vbRed
            ActiveCell = "P"
            ActiveCell.Font.Name = "Wingdings"
            Range("D1").Value = Range("D1").Value - 1
        'Removes flag if one already exists.
        Else
            ActiveCell.Font.Color = vbBlack
            ActiveCell.ClearContents
            ActiveCell.Font.Name = "Calibri"
            Range("D1").Value = Range("D1").Value + 1
        End If
    End If
    
    'Winning sequence for revealing all cells without bombs and correctly flagging all cells with bombs.
    If WorksheetFunction.Sum(WinField) = WorksheetFunction.Sum(Minefield) And Range("D1").Value = 0 Then
        SubsequentMove = False
        MsgBox ("Congratulations!!! You win!!!")
        Call Reset
    End If
Else
    MsgBox ("Select a cell within the minefield.")
End If

ActiveSheet.Protect

End Sub

Sub SeriousFace(Finish As Long)

'This sub causes a time delay that allows the viewer to perceive a serious face in J1 before it switches back to a happy face.


Dim EndTick As Integer
Dim Looper As Integer


EndTick = (Finish * 500)

For Looper = 1 To EndTick
    With Range("J1")
        .Font.Name = "Wingdings"
        .Value = "K"
    End With
Next Looper


End Sub

