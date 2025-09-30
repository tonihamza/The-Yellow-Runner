' Class module: BlueCell

Option Explicit

' Properties
Public blueCell As Range
Public DirectionCol As Integer
Public DirectionRow As Integer
Public BlackColor As Long
Public targetCell As Range

' Initialize the blue cell object
Public Sub Initialize(startCell As Range)
Set blueCell = startCell
BlackColor = RGB(0, 0, 0) ' Default black color


' Set initial directions (move diagonally down-right by default)
DirectionCol = 1
DirectionRow = 1
End Sub

' Move the blue cell diagonally
Public Sub Move()
Dim cellToMoveTo As Range

' Calculate the next cell based on current direction
On Error Resume Next
Set cellToMoveTo = blueCell.offset(DirectionRow, DirectionCol)
On Error GoTo 0

' Check if the next cell is within bounds and not black
If Not cellToMoveTo Is Nothing And Not (cellToMoveTo.Interior.color = BlackColor Or cellToMoveTo.offset(-DirectionRow, 0).Interior.color = BlackColor Or cellToMoveTo.offset(0, -DirectionCol).Interior.color = BlackColor) Then
    ' Move is valid
    RemoveBlueConditionalFormat blueCell ' Remove blue formatting from the current cell
    Set blueCell = cellToMoveTo
    ApplyBlueConditionalFormat blueCell ' Apply blue formatting to the new cell
Else
    ' Cell is black or out of bounds, change direction
    ChangeDirection
End If
End Sub

' Change direction based on surrounding cells
Private Sub ChangeDirection()
Dim cellToCheck As Range
Dim isBlackAhead As Boolean


' Check if the cell in the current direction is black
On Error Resume Next
Set cellToCheck = blueCell.offset(DirectionRow, DirectionCol)
On Error GoTo 0

isBlackAhead = Not cellToCheck Is Nothing And (cellToCheck.Interior.color = BlackColor Or cellToCheck.offset(-DirectionRow, 0).Interior.color = BlackColor Or cellToCheck.offset(0, -DirectionCol).Interior.color = BlackColor)

If isBlackAhead Then
    ' Current direction cell is black, determine which direction to reverse
    On Error Resume Next
    ' Check if cell to the right or left (same row) is black
    If cellToCheck.offset(-DirectionRow, 0).Interior.color <> BlackColor And cellToCheck.offset(0, -DirectionCol).Interior.color <> BlackColor Then
        DirectionCol = -DirectionCol
        DirectionRow = -DirectionRow
    End If
    Set cellToCheck = blueCell.offset(0, DirectionCol) ' Check right or left
    If Not cellToCheck Is Nothing And cellToCheck.Interior.color = BlackColor Then
        DirectionCol = -DirectionCol ' Reverse column direction
    End If
    
    ' Check if cell above or below (same column) is black
    Set cellToCheck = blueCell.offset(DirectionRow, 0) ' Check up or down
    If Not cellToCheck Is Nothing And cellToCheck.Interior.color = BlackColor Then
        DirectionRow = -DirectionRow ' Reverse row direction
    End If
    On Error GoTo 0
End If
End Sub

' ====================
' Apply blue conditional formatting to a cell
' ====================
Private Sub ApplyBlueConditionalFormat(cell As Range)
' Remove existing conditional formats to avoid duplicates
cell.FormatConditions.Delete
' Add conditional formatting to make the cell blue
With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
.Interior.color = RGB(0, 0, 255)
End With
End Sub

' ====================
' Remove blue conditional formatting from a cell
' ====================
Private Sub RemoveBlueConditionalFormat(cell As Range)
cell.FormatConditions.Delete
End Sub
