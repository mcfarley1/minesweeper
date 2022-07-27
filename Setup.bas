Attribute VB_Name = "Setup"
Sub Setup()
Attribute Setup.VB_ProcData.VB_Invoke_Func = " \n14"

ActiveSheet.Unprotect

'Sets alignment, font sizes, column widths, and row heights.
With Columns("A:U")
    .Font.Size = 16
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .ColumnWidth = 8.32
End With

With Columns("L")
    .Font.Size = 14
    .HorizontalAlignment = xlLeft
End With

Range("C1").HorizontalAlignment = xlRight

Rows("3:10").RowHeight = 45.6

'Gets rid of grid lines.
ActiveWindow.DisplayGridlines = False

'Create borders.
Range("C3:J10").Interior.Color = RGB(200, 200, 200)
Range("C3:J10").Borders(xlDiagonalDown).LineStyle = xlNone
Range("C3:J10").Borders(xlDiagonalUp).LineStyle = xlNone
With Range("C3:J10").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("C3:J10").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("C3:J10").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("C3:J10").Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("C3:J10").Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range("C3:J10").Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

'Add buttons.
ActiveSheet.Buttons.Add(226.5, 7, 96.3, 25.2).Select
Selection.OnAction = "Play"
Selection.Characters.Text = "Poke (Ctrl+Z)"

ActiveSheet.Buttons.Add(350.5, 7, 96.3, 25.2).Select
Selection.OnAction = "Flag"
Selection.Characters.Text = "Flag (Ctrl+X)"

ActiveSheet.Buttons.Add(726.5, 7, 96.3, 25.2).Select
Selection.OnAction = "Reset"
Selection.Characters.Text = "Reset (Ctrl+V)"

Range("F6").Select

'Add text.
Range("L3") = "1.  Poke a square to start the game by pressing the button or Ctrl+Z."
Range("L4") = "2.  The numbers indicate the number of mines in the surrounding cells."
Range("L5") = "3.  Use flags by pressing the button or Ctrl+X to denote cells with mines."
Range("L6") = "4.  The mine counter in cell D1 tells you how many mines remain to be found."
Range("L7") = "5.  When you have flagged all of the mines and cleared all of the remaining cells, you win!"
Range("L8") = "6.  The game can be reset at any time by pressing the button or Ctrl+V."
Range("C1") = "Mines remaining:"

'Sets the initial unfound bomb count to 10.
Range("D1") = 10

'Places a happy face in J1.
With Range("J1")
    .Font.Name = "Wingdings"
    .Value = "J"
End With

ActiveSheet.Protect

End Sub

