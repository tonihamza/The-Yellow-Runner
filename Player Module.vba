Option Explicit

#If Win64 Then
Public Declare PtrSafe Function SetTimer Lib "User32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As LongPtr, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "User32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Public TimerID As LongPtr
#Else
Public Declare Function SetTimer Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

Public TimerID As Long
#End If

' Movement direction variables
Public DirectionRow As Long
Public DirectionCol As Long
Public TempDirectionRow As Long ' Temporary direction row
Public TempDirectionCol As Long ' Temporary direction col
Public go As Boolean

' Constants
Const MoveIntervalSelection = 50 ' Time interval for each movement in milliseconds

' ====================
' Start the timer to initiate movement
' ====================
Sub StartTimerSelection()
    If TimerID <> 0 Then KillTimer 0, TimerID ' Kill any existing timer
    TimerID = SetTimer(0, 0, MoveIntervalSelection, AddressOf TimerEventSelection)
    go = True
End Sub

' ====================
' Stop the timer
' ====================
Public Sub StopTimerSelection()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
End Sub

' ====================
' Timer event calls the MoveSelection subroutine
' ====================
Sub TimerEventSelection()
    On Error Resume Next
    MoveSelection
End Sub

' ====================
' Movement logic for the selection
' ====================
Sub MoveSelection()
    On Error Resume Next
    Dim currentCell As Range
    Dim nextCell As Range
    Dim blackOrYellowCount As Long ' Count cells that are black or yellow

    Set currentCell = ActiveCell
    currentCell.Interior.color = RGB(255, 255, 0) ' Mark current cell as visited

    ' Check if we can apply the new direction (temp direction)
    If Not IsCellBlocked(currentCell.Offset(TempDirectionRow, TempDirectionCol)) Then
        DirectionRow = TempDirectionRow
        DirectionCol = TempDirectionCol
    End If
    
    ' Calculate the next cell based on the (updated) direction
    Set nextCell = currentCell.Offset(DirectionRow, DirectionCol)

    If Not IsCellBlocked(nextCell) Then
        nextCell.Select
    End If

    ' Check if next cell has any conditional formats and stop the game if it does
    If nextCell.FormatConditions.Count > 0 And go Then
        go = False
        Call StopAllTimers
        MsgBox "Game Over", vbInformation
        Call StartGame
    End If
    
    ' Count cells that are black or yellow
    blackOrYellowCount = CountBlackOrYellowCells()

    ' Level completion logic
    If blackOrYellowCount >= p And go Then
        go = False
        Call StopAllTimers
        level = level + 1
        MsgBox "Level Completed", vbInformation
        Call StartGame
    End If
End Sub

' ====================
' Check if a cell is blocked (black)
' ====================
Private Function IsCellBlocked(cell As Range) As Boolean
    IsCellBlocked = (cell.Interior.color = RGB(0, 0, 0)) ' Returns True if cell is black
End Function

' ====================
' Count cells that are black or yellow
' ====================
Private Function CountBlackOrYellowCells() As Long
    Dim cell As Range
    Dim count As Long
    count = 0
    
    ' Loop through a specific range to count cells
    For Each cell In Range("S4:AT31").Cells
        If cell.Interior.color = RGB(0, 0, 0) Or cell.Interior.color = RGB(255, 255, 0) Then
            count = count + 1
        End If
    Next cell

    CountBlackOrYellowCells = count
End Function

' ====================
' Stop all timers
' ====================
Private Sub StopAllTimers()
    Call StopTimerSelection
    Call StopTimerRed
    Call StopTimerBlue
    Call StopTimerPurple
End Sub

' ====================
' Movement in four directions
' ====================
Sub MoveUp()
    TempDirectionRow = -1
    TempDirectionCol = 0
End Sub

Sub MoveDown()
    TempDirectionRow = 1
    TempDirectionCol = 0
End Sub

Sub MoveLeft()
    TempDirectionRow = 0
    TempDirectionCol = -1
End Sub

Sub MoveRight()
    TempDirectionRow = 0
    TempDirectionCol = 1
End Sub

' ====================
' Start level and associated timers
' ====================
Sub StartLevel()
    Call StartTimerRed
    Call StartTimerBlue
    Call StartTimerPurple
    Call StartTimerBrown
    Call StartTimerMagenta
    StartTimerSelection
End Sub
