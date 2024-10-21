'Class Module: magentaCell

Option Explicit

' Properties
Private positionRange As Range ' Range of cells between the start and end positions
Private moveCounter As Integer ' Counter to keep track of moves
Private MagentaColor As Long ' Store the magenta color value
Private innerRange As Range ' Range excluding the first and last cells
Private startPos As Range ' First cell in the range
Private endPos As Range ' Last cell in the range

' ====================
' Initialize the class with start and end positions
' ====================
Public Sub Initialize(positions As Variant)
    ' Assume positions contains two values (start and end positions)
    Set startPos = Range(positions(0))
    Set endPos = Range(positions(1))

    ' Save the full range between start and end positions
    Set positionRange = Range(startPos, endPos)
    
    ' Exclude the first and last cells to form the inner range
    If positionRange.Cells.count > 2 Then
        If startPos.Column = endPos.Column Then
            Set innerRange = Range(startPos.offset(1, 0), endPos.offset(-1, 0))
        Else
            Set innerRange = Range(startPos.offset(0, 1), endPos.offset(0, -1))
        End If
    Else
        Set innerRange = Nothing ' No inner range if only two cells are involved
    End If
    startPos.Interior.color = RGB(0, 0, 0)
    endPos.Interior.color = RGB(0, 0, 0)
    moveCounter = 0 ' Initialize move counter
    MagentaColor = RGB(255, 0, 255) ' Set magenta color
    
    ' Always apply initial magenta formatting on the first and last cell
    ApplyMagentaConditionalFormat startPos
    ApplyMagentaConditionalFormat endPos
End Sub

' ====================
' Apply magenta formatting to the inner cells only (excluding the first and last cells)
' ====================
Private Sub ApplyMagentaEffectToInnerCells()
    Dim cell As Range
    
    ' Apply magenta effect to every cell in the inner range (if it exists)
    If Not innerRange Is Nothing Then
        For Each cell In innerRange
            ApplyMagentaConditionalFormat cell
        Next cell
    End If
End Sub

' ====================
' Remove magenta formatting from the inner cells only (excluding the first and last cells)
' ====================
Private Sub RemoveMagentaEffectFromInnerCells()
    Dim cell As Range
    
    ' Remove magenta effect from every cell in the inner range (if it exists)
    If Not innerRange Is Nothing Then
        For Each cell In innerRange
            RemoveMagentaConditionalFormat cell
        Next cell
    End If
End Sub

' ====================
' Apply magenta conditional formatting to a cell
' ====================
Private Sub ApplyMagentaConditionalFormat(cell As Range)
    ' Remove any existing conditional formats to avoid duplicates
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

' ====================
' Change the color of the first and last cells gradually using conditional formatting
' ====================
Private Sub UpdateFirstAndLastCellColors()
    Dim darkMagenta As Long
    Dim midMagenta As Long
    Dim lightMagenta As Long
    
    ' Define shades of magenta
    darkMagenta = RGB(139, 0, 139) ' Dark magenta
    midMagenta = RGB(200, 0, 200) ' Medium magenta
    lightMagenta = RGB(230, 0, 230) ' Light magenta

    ' Remove existing conditional formats to avoid conflicts
    startPos.FormatConditions.Delete
    endPos.FormatConditions.Delete

    ' Apply conditional formatting based on the moveCounter
    Select Case moveCounter
        Case 1
            ApplyConditionalFormattingWithColor startPos, darkMagenta
            ApplyConditionalFormattingWithColor endPos, darkMagenta
        Case 2
            ApplyConditionalFormattingWithColor startPos, midMagenta
            ApplyConditionalFormattingWithColor endPos, midMagenta
        Case 3
            ApplyConditionalFormattingWithColor startPos, lightMagenta
            ApplyConditionalFormattingWithColor endPos, lightMagenta
        Case 4
            ApplyConditionalFormattingWithColor startPos, MagentaColor ' Normal magenta
            ApplyConditionalFormattingWithColor endPos, MagentaColor ' Normal magenta
    End Select
End Sub

' ====================
' Apply conditional formatting to a cell with a specific color
' ====================
Private Sub ApplyConditionalFormattingWithColor(cell As Range, colorValue As Long)
    With cell.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.color = colorValue
    End With
End Sub

' ====================
' Move method for toggling magenta effect on/off for the inner cells and adjusting outer cells
' ====================
Public Sub Move()
    ' Increment the move counter
    moveCounter = moveCounter + 1

    ' Adjust the color of the first and last cells over the first 4 moves
    If moveCounter <= 4 Then
        UpdateFirstAndLastCellColors
    End If

    ' Apply the magenta effect to the inner cells every 5 moves
    If moveCounter = 5 Then
        ApplyMagentaEffectToInnerCells
    End If

    ' Remove the magenta effect from the inner cells every 6 moves and reset the counter
    If moveCounter > 6 Then
        moveCounter = 1 ' Reset the counter
        RemoveMagentaEffectFromInnerCells
    End If
End Sub

