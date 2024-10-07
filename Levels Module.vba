' Public variables
Public TimerIDRed As LongPtr
Public TimerIDBlue As LongPtr
Public TimerIDPurple As LongPtr ' New timer ID for purple cells
Public TimerIDBrown As LongPtr
Public RedCells As Collection ' Collection to hold all red cell objects
Public BlueCells As Collection ' Collection to hold all blue cell objects
Public PurpleCells As Collection ' Collection to hold all purple cell objects
Public BrownCells As Collection
Public MagentaCells As Collection
Public BlackColor As Long ' Default black color
Public level As Integer
Public MoveIntervalRed As Long ' Interval for red cell timer in milliseconds
Public MoveIntervalBlue As Long ' Interval for blue cell timer in milliseconds
Public MoveIntervalPurple As Long ' Interval for purple cell timer in milliseconds
Public MoveIntervalBrown As Long ' Interval for brown cell timer in milliseconds
Public TimerIDMagenta As LongPtr ' New timer ID for magenta cells
Public MoveIntervalMagenta As Long ' Interval for magenta cell timer in milliseconds

Public p As Integer

Public Sub StartGame()
    ' Define variables for red, blue, purple, and brown cell counts and positions
    Dim redPos As Variant
    Dim bluePos As Variant
    Dim purplePosH As Variant ' Horizontal purple positions
    Dim purplePosV As Variant ' Vertical purple positions
    Dim brownPos As Variant ' New positions for brown cells
    Dim magentaPos As Variant
    Dim i As Integer
    Dim rng As Range

    ' Set timer intervals (in milliseconds)
    MoveIntervalRed = 100 ' Example: 200 milliseconds
    MoveIntervalBlue = 100 ' Example: 200 milliseconds
    MoveIntervalPurple = 100 ' Example: 200 milliseconds
    MoveIntervalBrown = 100 ' Example: 200 milliseconds for brown cells
    MoveIntervalMagenta = 100

    If level = 0 Then level = 24
    Call ResetGame
    Range("AG31").Select

    ' Determine actions based on the current level
    Range("S4:AT31").Interior.color = RGB(255, 255, 255)
    ActiveSheet.Cells.FormatConditions.Delete
    p = 784
    Select Case level
        Case 1
            p = 784
            ' Set black color for specific ranges
            For i = 0 To 6
                Set rng = Range("T5:T30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i
            For i = 0 To 6
                Set rng = Range("AG5:AG30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i

        Case 2
            purplePosH = Array("AF17")
            purplePosV = Array("AF17")
            p = 584
        Case 3
            purplePosH = Array("S11", "AT23")
            purplePosV = Array("AA4", "AL31")
        Case 4
            purplePosV = Array("S11", "AT23", "AA4", "AL31")
            purplePosH = Array("AA4", "AL31", "S11", "AT23")
        Case 5
            redPos = Array("S4", "AT4")
        Case 6
            redPos = Array("S4", "AT4")
            Range("AB5:AB30").Interior.color = RGB(0, 0, 0)
            Range("AK5:AK30").Interior.color = RGB(0, 0, 0)
        Case 7
            bluePos = Array("S6", "AT6")
        Case 8
            redPos = Array("S4", "AT4")
            Range("U17:AR17").Interior.color = RGB(0, 0, 0)
            Range("AG6:AG29").Interior.color = RGB(0, 0, 0)
        Case 9
            Range("U17:AR17").Interior.color = RGB(0, 0, 0)
            Range("AG6:AG29").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9", "AT9", "S29", "AT29")
        Case 10
            Range("AC14:AJ21").Interior.color = RGB(0, 0, 0)
            purplePosV = Array("X4", "AH4")
            purplePosH = Array("S9", "S26")
            redPos = Array("S4")
        Case 11
            Range("AC14:AJ21").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9", "AT9", "S26", "AT26")
        Case 12
            For i = 0 To 6
                Set rng = Range("T5:T30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i
            For i = 0 To 6
                Set rng = Range("AG5:AG30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i
            redPos = Array("S4", "AT4")
        Case 13
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            purplePosH = Array("AF15", "AF7", "AF23", "AF28")
            purplePosV = Array("U10", "AB10", "AK10", "AR10")
        Case 14
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            redPos = Array("S4", "AT4")
        Case 15
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9")
        Case 16
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9", "AT9")
        Case 17
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S6", "S16")
        Case 18
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            redPos = Array("S4", "AT4", "S22", "AT22")
        Case 19
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9", "AT9", "S26", "AT26")
        Case 20
            redPos = Array("S4", "AT4")
            bluePos = Array("S6", "AT6")
            purplePosH = Array("S11", "AT23")
            purplePosV = Array("AA4", "AL31")
        Case 21
            ThisWorkbook.Sheets("Sheet1").Range("A130:AD159").Copy ThisWorkbook.Sheets("Sheet1").Range("R3:AU32")
            redPos = Array("AF15")
            MoveIntervalRed = 150
        Case 22
            ThisWorkbook.Sheets("Sheet1").Range("A130:AD159").Copy ThisWorkbook.Sheets("Sheet1").Range("R3:AU32")
            redPos = Array("AI15", "AC15")
            MoveIntervalRed = 150
        Case 23
            ThisWorkbook.Sheets("Sheet1").Range("A130:AD159").Copy ThisWorkbook.Sheets("Sheet1").Range("R3:AU32")
            redPos = Array("AI15", "AF15", "AC15")
            MoveIntervalRed = 150
        Case 24
            ThisWorkbook.Sheets("Sheet1").Range("A130:AD159").Copy ThisWorkbook.Sheets("Sheet1").Range("R3:AU32")
            redPos = Array("S4", "AS4", "S31", "AS31")
            MoveIntervalRed = 150
            Range("AF17").Select
        Case 25
            purplePosH = Array("S9", "S26")
            MoveIntervalPurple = 10
        Case 26
            bluePos = Array("S9", "S26")
            MoveIntervalBlue = 30
        Case 27
            purplePosH = Array("S4", "S6", "S8", "S10", "S12", "S14", "S16", "S18", "S20", "S22", "S24", "S26", "S28", "S30", "AT5", "AT7", "AT9", "AT11", "AT13", "AT15", "AT17", "AT19", "AT21", "AT23", "AT25", "AT27", "AT29", "AT31")
            MoveIntervalPurple = 400
        Case Else
            MsgBox "You finished the game!", vbExclamation
            Exit Sub
    End Select

    ' Call GenerateEnemy to initialize enemies based on the current level
    GenerateEnemy redPos, bluePos, purplePosH, purplePosV, brownPos, magentaPos
End Sub
Sub ResetGame()
    ' Reseteaza variabilele
    DirectionRow = 0
    DirectionCol = 0
    go = False
    
    ' Opre?te toate timerele
    Call StopTimerSelection
    Call StopTimerRed
    Call StopTimerBlue
    Call StopTimerPurple
    Call StopTimerBrown ' Stop timer for brown cells
    
    Range("S4:AT31").Interior.color = RGB(255, 255, 255)
    ActiveSheet.Cells.FormatConditions.Delete
    ' Resetarea culorilor celulelor sau orice alta logica necesara
    ' De exemplu:
    ' Dim cell As Range
    ' For Each cell In Range("S4:AT31").Cells
    '     cell.Interior.Color = RGB(255, 255, 255) ' Schimba culoarea la alb sau alta culoare implicita
    ' Next cell
End Sub


' Initialize the red, blue, and purple cell movements
Sub GenerateEnemy(redPositions As Variant, bluePositions As Variant, purplePositionsH As Variant, purplePositionsV As Variant, brownPositions As Variant, magentaPositions As Variant)
    Dim redCell As redCell
    Dim blueCell As blueCell
    Dim purpleCellH As PurpleCell ' Horizontal purple cell
    Dim purpleCellV As PurpleCell ' Vertical purple cell
    Dim brownCell As brownCell ' Brown cell
    Dim i As Integer
    Dim targetCell1 As Range
    
    ' Set the initial target cell to the current selection
    Set targetCell1 = ActiveCell
    
    ' Clear previous collections if they exist
    On Error Resume Next
    Set RedCells = Nothing
    Set BlueCells = Nothing
    Set PurpleCells = Nothing
    Set BrownCells = Nothing
    Set MagentaCells = Nothing ' Initialize brown cell collection
    On Error GoTo 0
    
    Set RedCells = New Collection
    Set BlueCells = New Collection
    Set PurpleCells = New Collection
    Set BrownCells = New Collection
    Set MagentaCells = New Collection
    
    ' Check if redPositions array is not empty and initialize red cell objects
    If Not IsEmpty(redPositions) Then
        For i = LBound(redPositions) To UBound(redPositions)
        
                Set redCell = New redCell
                redCell.Initialize Range(redPositions(i)), i Mod 2 = 0 ' Example: alternate movement directions
                RedCells.Add redCell
                With Range(redPositions(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(255, 0, 0)
                End With
        Next i
    End If
    
    ' Check if bluePositions array is not empty and initialize blue cell objects
    If Not IsEmpty(bluePositions) Then
        For i = LBound(bluePositions) To UBound(bluePositions)
                Set blueCell = New blueCell
                blueCell.Initialize Range(bluePositions(i))
                BlueCells.Add blueCell
                With Range(bluePositions(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(0, 0, 255)
                End With
        Next i
    End If
    
    ' Check if purplePositionsH array is not empty and initialize horizontal purple cell objects
    If Not IsEmpty(purplePositionsH) Then
        For i = LBound(purplePositionsH) To UBound(purplePositionsH)
                Set purpleCellH = New PurpleCell
                purpleCellH.Initialize Range(purplePositionsH(i)), False ' Horizontal movement
                PurpleCells.Add purpleCellH
                With Range(purplePositionsH(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(128, 0, 128)
                End With
        Next i
    End If

    ' Check if purplePositionsV array is not empty and initialize vertical purple cell objects
    If Not IsEmpty(purplePositionsV) Then
        For i = LBound(purplePositionsV) To UBound(purplePositionsV)
                Set purpleCellV = New PurpleCell
                purpleCellV.Initialize Range(purplePositionsV(i)), True ' Vertical movement
                PurpleCells.Add purpleCellV
                With Range(purplePositionsV(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(128, 0, 128)
                End With
        Next i
    End If
        ' Check if brownPositions array is not empty and initialize brown cell objects
    If Not IsEmpty(brownPositions) Then
        For i = LBound(brownPositions) To UBound(brownPositions)
                Set brownCell = New brownCell
                brownCell.Initialize Range(brownPositions(i))
                BrownCells.Add brownCell
                With Range(brownPositions(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(165, 42, 42) ' Brown color
                End With
        Next i
    End If
    ' Check if magentaPositions array is not empty and initialize magenta cell objects
If Not IsEmpty(magentaPositions) Then
    For i = LBound(magentaPositions) To UBound(magentaPositions)
        Dim magentaCell As magentaCell
        Set magentaCell = New magentaCell
        magentaCell.Initialize magentaPositions(i) ' Initialize magenta cell with string
        MagentaCells.Add magentaCell
    Next i
End If

End Sub


' Timer handling and movement logic for red cells
Public Sub StartTimerRed()
    If TimerIDRed <> 0 Then KillTimer 0, TimerIDRed
    TimerIDRed = SetTimer(0, 0, MoveIntervalRed, AddressOf TimerEventRed)
End Sub

Public Sub StopTimerRed()
    If TimerIDRed <> 0 Then KillTimer 0, TimerIDRed
End Sub

Sub TimerEventRed()
    On Error Resume Next
    Call MoveRedCells
End Sub

Sub MoveRedCells()
    Dim redCell As redCell
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To RedCells.count
        Set redCell = RedCells(i)
        redCell.Move ActiveCell
        If redCell.redCell.Address = ActiveCell.Address Then anyTargetReached = True
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

' Timer handling and movement logic for blue cells
Public Sub StartTimerBlue()
    If TimerIDBlue <> 0 Then KillTimer 0, TimerIDBlue
    TimerIDBlue = SetTimer(0, 0, MoveIntervalBlue, AddressOf TimerEventBlue)
End Sub

Public Sub StopTimerBlue()
    If TimerIDBlue <> 0 Then KillTimer 0, TimerIDBlue
End Sub

Sub TimerEventBlue()
    On Error Resume Next
    Call MoveBlueCells
End Sub

Sub MoveBlueCells()
    Dim blueCell As blueCell
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To BlueCells.count
        Set blueCell = BlueCells(i)
        blueCell.Move
        If blueCell.blueCell.Address = ActiveCell.Address Then anyTargetReached = True
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

' Timer handling and movement logic for purple cells
Public Sub StartTimerPurple()
    If TimerIDPurple <> 0 Then KillTimer 0, TimerIDPurple
    TimerIDPurple = SetTimer(0, 0, MoveIntervalPurple, AddressOf TimerEventPurple)
End Sub

Public Sub StopTimerPurple()
    If TimerIDPurple <> 0 Then KillTimer 0, TimerIDPurple
End Sub

Sub TimerEventPurple()
    On Error Resume Next
    Call MovePurpleCells
End Sub

Sub MovePurpleCells()
    Dim PurpleCell As PurpleCell
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To PurpleCells.count
        Set PurpleCell = PurpleCells(i)
        PurpleCell.Move
        If PurpleCell.PurpleCell.Address = ActiveCell.Address Then anyTargetReached = True
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

' Timer handling and movement logic for brown cells
Public Sub StartTimerBrown()
    If TimerIDBrown <> 0 Then KillTimer 0, TimerIDBrown
    TimerIDBrown = SetTimer(0, 0, MoveIntervalBrown, AddressOf TimerEventBrown)
End Sub

Public Sub StopTimerBrown()
    If TimerIDBrown <> 0 Then KillTimer 0, TimerIDBrown
End Sub

Sub TimerEventBrown()
    On Error Resume Next
    Call MoveBrownCells
End Sub

Sub MoveBrownCells()
    Dim brownCell As brownCell
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To BrownCells.count
        Set brownCell = BrownCells(i)
        brownCell.Move
        If brownCell.brownCell.Address = ActiveCell.Address Then anyTargetReached = True
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub
' Timer for magenta cells
Public Sub StartTimerMagenta()
    TimerIDMagenta = SetTimer(0, 0, MoveIntervalMagenta, AddressOf TimerMagentaHandler)
End Sub

' Stop timer for magenta cells
Public Sub StopTimerMagenta()
    If TimerIDMagenta <> 0 Then
        KillTimer 0, TimerIDMagenta
        TimerIDMagenta = 0
    End If
End Sub

' Handler for magenta cell movements
Public Sub TimerMagentaHandler()
    Dim magentaCell As magentaCell
    Dim i As Integer
    For i = 1 To BrownCells.count
        Set magentaCell = MagentaCells(i)
        magentaCell.Move
    Next i
End Sub


' Stop all timers
Public Sub StopAllTimers()
    StopTimerRed
    StopTimerBlue
    StopTimerPurple
    StopTimerBrown
    Call StopTimerSelection
End Sub
