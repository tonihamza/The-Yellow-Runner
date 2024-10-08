' Class module: RedCell

Option Explicit

' Properties
Public redCell As Range
Public DirectionKKK As Integer
Public DirectionLLL As Integer
Public BlackColor As Long
Public RowReached As Boolean
Public ColReached As Boolean

' ====================
' Initialize the red cell object
' ====================
Public Sub Initialize(startCell As Range, Optional startRowReached As Boolean = False)
    Set redCell = startCell
    BlackColor = RGB(0, 0, 0) ' Default black color
    RowReached = startRowReached
    ColReached = False

    ' Set directions based on RowReached status
    If RowReached Then
        DirectionKKK = 1 ' Default direction to move right
        DirectionLLL = -1 ' Default direction to move up
    Else
        DirectionKKK = -1 ' Default direction to move left
        DirectionLLL = 1 ' Default direction to move down
    End If
End Sub

' ====================
' Move the red cell towards its target cell
' ====================
Public Sub Move(TargetCell As Range)
    Dim cellToMoveTo As Range

    ' Check if the red cell has not reached the target cell
    If redCell.Address <> TargetCell.Address Then

        ' Move row-wise first
        If Not RowReached Then
            ' Check if the red cell's row is not equal to the target row
            If redCell.Row <> TargetCell.Row Then
                ' Determine movement direction
                If redCell.Row < TargetCell.Row Then
                    Set cellToMoveTo = redCell.Offset(1, 0) ' Move Down
                Else
                    Set cellToMoveTo = redCell.Offset(-1, 0) ' Move Up
                End If
                
                ' Check if moving row-wise is blocked
                If Not cellToMoveTo Is Nothing And cellToMoveTo.Interior.color <> BlackColor Then
                    ' Move is valid
                    RemoveRedConditionalFormat redCell ' Remove red formatting from the current cell
                    Set redCell = cellToMoveTo ' Update red cell position
                    ApplyRedConditionalFormat redCell ' Apply red formatting to the new cell
                Else
                    ' Row-wise movement blocked, try moving in direction kkk
                    Set cellToMoveTo = redCell.Offset(0, DirectionKKK)
                    
                    ' Check if movement in direction kkk is valid
                    If Not cellToMoveTo Is Nothing And cellToMoveTo.Interior.color <> BlackColor Then
                        ' Move is valid
                        RemoveRedConditionalFormat redCell ' Remove red formatting from the current cell
                        Set redCell = cellToMoveTo ' Update red cell position
                        ApplyRedConditionalFormat redCell ' Apply red formatting to the new cell
                    Else
                        ' Direction kkk blocked, invert direction
                        DirectionKKK = -DirectionKKK
                    End If
                End If
            Else
                ' Row has been reached
                RowReached = True
                ColReached = False
            End If
        End If
        
        ' If row is reached, start moving column-wise
        If RowReached And Not ColReached Then
            ' Check if the red cell's column is not equal to the target column
            If redCell.Column <> TargetCell.Column Then
                ' Determine movement direction
                If redCell.Column < TargetCell.Column Then
                    Set cellToMoveTo = redCell.Offset(0, 1) ' Move Right
                Else
                    Set cellToMoveTo = redCell.Offset(0, -1) ' Move Left
                End If
                
                ' Check if moving column-wise is blocked
                If Not cellToMoveTo Is Nothing And cellToMoveTo.Interior.color <> BlackColor Then
                    ' Move is valid
                    RemoveRedConditionalFormat redCell ' Remove red formatting from the current cell
                    Set redCell = cellToMoveTo ' Update red cell position
                    ApplyRedConditionalFormat redCell ' Apply red formatting to the new cell
                Else
                    ' Column movement blocked, try moving in direction lll
                    Set cellToMoveTo = redCell.Offset(DirectionLLL, 0)
                    
                    ' Check if movement in direction lll is valid
                    If Not cellToMoveTo Is Nothing And cellToMoveTo.Interior.color <> BlackColor Then
                        ' Move is valid
                        RemoveRedConditionalFormat redCell ' Remove red formatting from the current cell
                        Set redCell = cellToMoveTo ' Update red cell position
                        ApplyRedConditionalFormat redCell ' Apply red formatting to the new cell
                    Else
                        ' Direction lll blocked, invert direction
                        DirectionLLL = -DirectionLLL
                    End If
                End If
            Else
                ' Column has been reached
                ColReached = True
                RowReached = False
            End If
        End If
    End If
End Sub

' ====================
' Apply red conditional formatting to a cell
' ====================
Private Sub ApplyRedConditionalFormat(cell As Range)
    ' Remove existing conditional formats to avoid duplicates
    cell.FormatConditions.Delete
    ' Add conditional formatting to make the cell red
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = RGB(255, 0, 0) ' Red color
    End With
End Sub

' ====================
' Remove red conditional formatting from a cell
' ====================
Private Sub RemoveRedConditionalFormat(cell As Range)
    ' Remove conditional formatting from the specified cell
    cell.FormatConditions.Delete
End Sub
