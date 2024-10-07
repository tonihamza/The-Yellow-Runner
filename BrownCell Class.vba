' Class module: BrownCell

Option Explicit

' Properties
Public brownCell As Range
Public DirectionCol As Integer
Public DirectionRow As Integer
Public BlackColor As Long
Public moveCounter As Integer ' Count number of consecutive moves in the same direction

' ====================
' Initialize the brown cell object
' ====================
Public Sub Initialize(startCell As Range)
    Set brownCell = startCell
    BlackColor = RGB(0, 0, 0) ' Default black color

    ' Set initial direction (move diagonally down-right by default)
    DirectionCol = 1
    DirectionRow = 1
    moveCounter = 0 ' Initialize move counter
End Sub

' ====================
' Move the brown cell in one of the 8 possible directions
' ====================
Public Sub Move()
    Dim cellToMoveTo As Range

    ' Calculate the next cell based on current direction
    On Error Resume Next
    Set cellToMoveTo = brownCell.Offset(DirectionRow, DirectionCol)
    On Error GoTo 0

    ' Check if the next cell is within bounds and not black
    If Not cellToMoveTo Is Nothing And Not (cellToMoveTo.Interior.color = BlackColor) Then
        ' Move is valid
        RemoveBrownConditionalFormat brownCell ' Remove brown formatting from the current cell
        Set brownCell = cellToMoveTo
        ApplyBrownConditionalFormat brownCell ' Apply brown formatting to the new cell
        moveCounter = moveCounter + 1 ' Increment the move counter
    Else
        ' Cell is black or out of bounds, change direction
        ChangeDirection
    End If

    ' Change direction if the move counter reaches 10
    If moveCounter >= 10 Then
        ChangeDirection
    End If
End Sub

' ====================
' Change direction based on surrounding cells
' ====================
Private Sub ChangeDirection()
    Dim directionIndex As Integer
    directionIndex = Int(8 * Rnd + 1) ' Generate a random number between 1 and 8

    Select Case directionIndex
        Case 1
            ' Up
            DirectionRow = -1
            DirectionCol = 0
        Case 2
            ' Down
            DirectionRow = 1
            DirectionCol = 0
        Case 3
            ' Left
            DirectionRow = 0
            DirectionCol = -1
        Case 4
            ' Right
            DirectionRow = 0
            DirectionCol = 1
        Case 5
            ' Diagonal up-left
            DirectionRow = -1
            DirectionCol = -1
        Case 6
            ' Diagonal up-right
            DirectionRow = -1
            DirectionCol = 1
        Case 7
            ' Diagonal down-left
            DirectionRow = 1
            DirectionCol = -1
        Case 8
            ' Diagonal down-right
            DirectionRow = 1
            DirectionCol = 1
    End Select

    moveCounter = 0 ' Reset the move counter after changing direction
End Sub

' ====================
' Apply brown conditional formatting to a cell
' ====================
Private Sub ApplyBrownConditionalFormat(cell As Range)
    ' Remove existing conditional formats to avoid duplicates
    cell.FormatConditions.Delete
    ' Add conditional formatting to make the cell brown
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = RGB(139, 69, 19) ' Brown color
    End With
End Sub

' ====================
' Remove brown conditional formatting from a cell
' ====================
Private Sub RemoveBrownConditionalFormat(cell As Range)
    cell.FormatConditions.Delete
End Sub
