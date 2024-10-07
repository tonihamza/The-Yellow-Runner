' Class module: MagentaCell

Option Explicit

' Properties
Private positionRange As Variant ' Array of positions to apply magenta effect
Private moveCounter As Integer ' Count number of moves
Private currentIndex As Integer ' Current index for position pairs
Private MagentaColor As Long ' Store the magenta color value
Private applyEffect As Boolean ' Flag to switch between applying and removing magenta effect

' ====================
' Initialize the class with a set of positions
' ====================
Public Sub Initialize(positions As Variant)
    Dim i As Integer
    positionRange = positions ' Store the positions (e.g., {"A3", "A8", "M8"})
    moveCounter = 0 ' Initialize move counter
    currentIndex = 0 ' Start with the first position
    MagentaColor = RGB(255, 0, 255) ' Magenta color
    applyEffect = True ' Start by applying the effect
    
    ' Apply initial magenta formatting to all specified positions
    For i = LBound(positions) To UBound(positions)
        With Range(positions(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
            .Interior.color = RGB(255, 0, 255) ' Magenta color
        End With
    Next i
End Sub

' ====================
' Move to the next position and apply magenta effect
' ====================
Public Sub Move()
    ' Remove the effect from current positions
    RemoveMagentaEffectBetweenPositions

    ' Increment currentIndex to move to the next pair of positions
    currentIndex = currentIndex + 1
    If currentIndex > UBound(positionRange) Then
        currentIndex = 1 ' Loop back to the first position
    End If
    
    ' Apply the effect between new positions
    ApplyMagentaEffectBetweenPositions
End Sub

' ====================
' Apply magenta formatting between two consecutive positions
' ====================
Private Sub ApplyMagentaEffectBetweenPositions()
    Dim startPos As Range
    Dim endPos As Range
    Dim cell As Range

    ' Get the start position from the positionRange array
    Set startPos = Range(positionRange(currentIndex))
    
    ' Check if the next index is within bounds
    If currentIndex + 1 <= UBound(positionRange) Then
        Set endPos = Range(positionRange(currentIndex + 1))

        ' Check if the cells are on the same row (horizontal) or the same column (vertical)
        If startPos.Row = endPos.Row Then
            ' Apply effect horizontally between startPos and endPos
            For Each cell In Range(startPos, endPos)
                ApplyMagentaConditionalFormat cell
            Next cell
        ElseIf startPos.Column = endPos.Column Then
            ' Apply effect vertically between startPos and endPos
            For Each cell In Range(startPos, endPos)
                ApplyMagentaConditionalFormat cell
            Next cell
        End If
    End If
End Sub

' ====================
' Remove magenta formatting between two consecutive positions
' ====================
Private Sub RemoveMagentaEffectBetweenPositions()
    Dim startPos As Range
    Dim endPos As Range
    Dim cell As Range

    ' Get the start position from the positionRange array
    Set startPos = Range(positionRange(currentIndex))
    
    ' Check if the next index is within bounds
    If currentIndex + 1 <= UBound(positionRange) Then
        Set endPos = Range(positionRange(currentIndex + 1))

        ' Check if the cells are on the same row (horizontal) or the same column (vertical)
        If startPos.Row = endPos.Row Then
            ' Remove effect horizontally between startPos and endPos
            For Each cell In Range(startPos, endPos)
                RemoveMagentaConditionalFormat cell
            Next cell
        ElseIf startPos.Column = endPos.Column Then
            ' Remove effect vertically between startPos and endPos
            For Each cell In Range(startPos, endPos)
                RemoveMagentaConditionalFormat cell
            Next cell
        End If
    End If
End Sub

' ====================
' Apply magenta conditional formatting to a cell
' ====================
Private Sub ApplyMagentaConditionalFormat(cell As Range)
    ' Remove existing conditional formats to avoid duplicates
    cell.FormatConditions.Delete
    ' Add conditional formatting to make the cell magenta
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = MagentaColor ' Apply magenta color
    End With
End Sub

' ====================
' Remove magenta conditional formatting from a cell
' ====================
Private Sub RemoveMagentaConditionalFormat(cell As Range)
    ' Remove all conditional formats from the cell
    cell.FormatConditions.Delete
End Sub
