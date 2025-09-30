' Class module: PurpleCell
Option Explicit

' Properties
Public PurpleCell As Range
Public DirectionCol As Integer
Public DirectionRow As Integer
Public BlackColor As Long

' Initialize the purple cell object
Public Sub Initialize(startCell As Range, moveVertically As Boolean)
    Set PurpleCell = startCell
    BlackColor = RGB(0, 0, 0) ' Default black color
    DirectionCol = IIf(moveVertically, 0, 1)
    DirectionRow = IIf(moveVertically, 1, 0)
End Sub

' Move the purple cell
Public Sub Move()
    Dim cellToMoveTo As Range
    Set cellToMoveTo = PurpleCell.offset(DirectionRow, DirectionCol)

    If Not cellToMoveTo Is Nothing And Not IsBlocked(cellToMoveTo) Then
        RemovePurpleConditionalFormat PurpleCell
        Set PurpleCell = cellToMoveTo
        ApplyPurpleConditionalFormat PurpleCell
    Else
        ChangeDirection
    End If
End Sub

' Change direction when blocked
Private Sub ChangeDirection()
    If IsBlocked(PurpleCell.offset(DirectionRow, DirectionCol)) Then
        DirectionRow = -DirectionRow
        DirectionCol = -DirectionCol
    End If
End Sub

' Check if a cell is blocked
Private Function IsBlocked(cell As Range) As Boolean
    IsBlocked = (cell.Interior.color = BlackColor)
End Function

' Apply purple conditional formatting to a cell
Private Sub ApplyPurpleConditionalFormat(cell As Range)
    cell.FormatConditions.Delete
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = RGB(128, 0, 128) ' Purple color
    End With
End Sub

' Remove purple conditional formatting from a cell
Private Sub RemovePurpleConditionalFormat(cell As Range)
    cell.FormatConditions.Delete
End Sub

