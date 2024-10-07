' Class module: PurpleCell

Option Explicit

' Properties
Public PurpleCell As Range
Public DirectionCol As Integer
Public DirectionRow As Integer
Public BlackColor As Long

' ====================
' Initialize the purple cell object
' ====================
Public Sub Initialize(startCell As Range, moveVertically As Boolean)
    Set PurpleCell = startCell
    BlackColor = RGB(0, 0, 0) ' Default black color

    ' Set initial direction based on the moveVertically parameter
    DirectionCol = IIf(moveVertically, 0, 1) ' Set column direction
    DirectionRow = IIf(moveVertically, 1, 0) ' Set row direction
End Sub

' ====================
' Move the purple cell
' ====================
Public Sub Move()
    Dim cellToMoveTo As Range
    Set cellToMoveTo = PurpleCell.Offset(DirectionRow, DirectionCol) ' Calculate the next cell

    ' Check if the next cell is within bounds and not blocked
    If Not cellToMoveTo Is Nothing And Not IsBlocked(cellToMoveTo) Then
        RemovePurpleConditionalFormat PurpleCell ' Remove formatting from current cell
        Set PurpleCell = cellToMoveTo ' Move to the new cell
        ApplyPurpleConditionalFormat PurpleCell ' Apply formatting to the new cell
    Else
        ChangeDirection ' Change direction if the next cell is blocked
    End If
End Sub

' ====================
' Change direction when blocked
' ====================
Private Sub ChangeDirection()
    ' Reverse direction if the next cell is blocked
    If IsBlocked(PurpleCell.Offset(DirectionRow, DirectionCol)) Then
        DirectionRow = -DirectionRow ' Reverse row direction
        DirectionCol = -DirectionCol ' Reverse column direction
    End If
End Sub

' ====================
' Check if a cell is blocked
' ====================
Private Function IsBlocked(cell As Range) As Boolean
    ' Return true if the cell is black
    IsBlocked = (cell.Interior.color = BlackColor)
End Function

' ====================
' Apply purple conditional formatting to a cell
' ====================
Private Sub ApplyPurpleConditionalFormat(cell As Range)
    ' Remove existing conditional formats to avoid duplicates
    cell.FormatConditions.Delete
    ' Add conditional formatting to make the cell purple
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = RGB(128, 0, 128) ' Purple color
    End With
End Sub

' ====================
' Remove purple conditional formatting from a cell
' ====================
Private Sub RemovePurpleConditionalFormat(cell As Range)
    ' Remove conditional formatting from the specified cell
    cell.FormatConditions.Delete
End Sub
