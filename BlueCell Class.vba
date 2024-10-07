' Class module: BlueCell

Option Explicit

' Properties
Public blueCell As Range ' The cell representing the blue entity in the game
Public DirectionCol As Integer ' The column direction for movement
Public DirectionRow As Integer ' The row direction for movement
Public BlackColor As Long ' Color value representing black
Public TargetCell As Range ' The target cell for the blue entity (if needed)

' ====================
' Initialize the blue cell object
' ====================
Public Sub Initialize(startCell As Range)
    ' Set the starting cell for the blue entity
    Set blueCell = startCell
    BlackColor = RGB(0, 0, 0) ' Default black color
    
    ' Set initial directions (move diagonally down-right by default)
    DirectionCol = 1 ' Move right in the column
    DirectionRow = 1 ' Move down in the row
End Sub

' ====================
' Move the blue cell diagonally
' ====================
Public Sub Move()
    Dim cellToMoveTo As Range ' The next cell to move to

    ' Calculate the next cell based on current direction
    On Error Resume Next ' Ignore errors temporarily
    Set cellToMoveTo = blueCell.offset(DirectionRow, DirectionCol) ' Calculate next cell position
    On Error GoTo 0 ' Restore normal error handling

    ' Check if the next cell is within bounds and not black
    If Not cellToMoveTo Is Nothing And Not (cellToMoveTo.Interior.color = BlackColor Or _
        cellToMoveTo.offset(-DirectionRow, 0).Interior.color = BlackColor Or _
        cellToMoveTo.offset(0, -DirectionCol).Interior.color = BlackColor) Then
        ' Move is valid
        RemoveBlueConditionalFormat blueCell ' Remove blue formatting from the current cell
        Set blueCell = cellToMoveTo ' Update the blue cell position
        ApplyBlueConditionalFormat blueCell ' Apply blue formatting to the new cell
    Else
        ' Cell is black or out of bounds, change direction
        ChangeDirection ' Call function to change the direction of movement
    End If
End Sub

' ====================
' Change direction based on surrounding cells
' ====================
Private Sub ChangeDirection()
    Dim cellToCheck As Range ' Cell to check for obstacles
    Dim isBlackAhead As Boolean ' Flag to check if there is a black cell ahead

    ' Check if the cell in the current direction is black
    On Error Resume Next ' Ignore errors temporarily
    Set cellToCheck = blueCell.offset(DirectionRow, DirectionCol) ' Get cell in the current direction
    On Error GoTo 0 ' Restore normal error handling

    ' Determine if there is a black cell in the direction the blue cell is moving
    isBlackAhead = Not cellToCheck Is Nothing And (cellToCheck.Interior.color = BlackColor Or _
        cellToCheck.offset(-DirectionRow, 0).Interior.color = BlackColor Or _
        cellToCheck.offset(0, -DirectionCol).Interior.color = BlackColor)

    If isBlackAhead Then
        ' Current direction cell is black, determine which direction to reverse
        On Error Resume Next ' Ignore errors temporarily
        ' Check if cell to the right or left (same row) is black
        If cellToCheck.offset(-DirectionRow, 0).Interior.color <> BlackColor And _
           cellToCheck.offset(0, -DirectionCol).Interior.color <> BlackColor Then
            DirectionCol = -DirectionCol ' Reverse column direction
            DirectionRow = -DirectionRow ' Reverse row direction
        End If
        
        ' Check the cell to the right or left (same row)
        Set cellToCheck = blueCell.offset(0, DirectionCol) ' Check right or left
        If Not cellToCheck Is Nothing And cellToCheck.Interior.color = BlackColor Then
            DirectionCol = -DirectionCol ' Reverse column direction
        End If
        
        ' Check if cell above or below (same column) is black
        Set cellToCheck = blueCell.offset(DirectionRow, 0) ' Check up or down
        If Not cellToCheck Is Nothing And cellToCheck.Interior.color = BlackColor Then
            DirectionRow = -DirectionRow ' Reverse row direction
        End If
        On Error GoTo 0 ' Restore normal error handling
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
        .Interior.color = RGB(0, 0, 255) ' Set the interior color to blue
    End With
End Sub

' ====================
' Remove blue conditional formatting from a cell
' ====================
Private Sub RemoveBlueConditionalFormat(cell As Range)
    ' Remove all conditional formatting from the specified cell
    cell.FormatConditions.Delete
End Sub
