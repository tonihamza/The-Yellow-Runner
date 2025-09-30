' Modulul Levels:
' Public variables
Public TimerIDRed As LongPtr
Public TimerIDBlue As LongPtr
Public TimerIDPurple As LongPtr ' Timer ID for purple cells
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
Public TimerIDSnake As LongPtr
Public MoveIntervalSnake As Long
Public Snakes As Collection
Public TimerIDBigRed As LongPtr
Public MoveIntervalBigRed As Long
Public BigRedEnemies As Collection

'--- Variabile pentru PurpleCellv2 ---
Public TimerIDPurpleV2 As LongPtr
Public MoveIntervalPurpleV2 As Long
Public PurpleCellsV2 As Collection
Public purplePosV2 As Variant

'--- Variabile pentru SpaceInvader ---
Public TimerIDSpaceInvader As LongPtr
Public MoveIntervalSpaceInvader As Long
Public SpaceInvaders As Collection

Public p As Integer

Public Sub StartGame()
    ' Define variables for red, blue, purple, and brown cell counts and positions
    Dim redPos As Variant
    Dim bluePos As Variant
    Dim purplePosH As Variant
    Dim purplePosV As Variant
    Dim brownPos As Variant
    Dim magentaPos As Variant
    Dim snakePos As Variant
    Dim bigRedPos As Variant
    Dim invaderPos As Variant
    
    Dim i As Integer
    Dim rng As Range

    ' Set timer intervals (in milliseconds)
    MoveIntervalRed = 100 ' Example: 200 milliseconds
    MoveIntervalBlue = 100 ' Example: 200 milliseconds
    MoveIntervalPurple = 100 ' Example: 200 milliseconds
    MoveIntervalBrown = 100 ' Example: 200 milliseconds for brown cells
    MoveIntervalMagenta = 300
    MoveIntervalPurpleV2 = 100
    MoveIntervalBigRed = 100
    MoveIntervalSpaceInvader = 150

    If level = 0 Then level = 1
    Call ResetGame
    Range("AG31").Select

    ' Determine actions based on the current level
    Range("S4:AT31").Interior.color = RGB(255, 255, 255)
    ActiveSheet.Cells.FormatConditions.Delete
    p = 600
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
            purplePosV = Array("AF17")
            purplePosH = Array("AF17")
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
        Case 28
            magentaPos = Array(Array("U6", "U29"), Array("U29", "AR29"), Array("U6", "AR6"), Array("AR6", "AR29"), Array("Y10", "Y25"), Array("Y25", "AN25"), Array("Y10", "AN10"), Array("AN10", "AN25"))
        Case 29
            brownPos = Array("S9", "AT9")
        Case 30
            brownPos = Array("S9", "S26", "AT9", "AT26")
        Case 31
             magentaPos = Array(Array("Y3", "Y32"), Array("AN3", "AN32"), Array("R10", "AU10"), Array("R25", "AU25"))
             redPos = Array("S4", "AT4")
        Case 32
            magentaPos = Array(Array("U6", "U29"), Array("U29", "AR29"), Array("U6", "AR6"), Array("AR6", "AR29"), Array("Y10", "Y25"), Array("Y25", "AN25"), Array("Y10", "AN10"), Array("AN10", "AN25"))
            bluePos = Array("S9", "S26")
        Case 33
            For i = 0 To 6
                Set rng = Range("T5:T30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i
            For i = 0 To 6
                Set rng = Range("AG5:AG30").offset(0, 2 * i)
                rng.Interior.color = RGB(0, 0, 0)
            Next i
            magentaPos = Array(Array("R17", "AU17"), Array("R10", "AU10"), Array("R25", "AU25"))
        Case 34
            purplePosV2 = Array("S9")
        Case 35
            magentaPos = Array( _
                Array("R4", "AU4"), _
                Array("R6", "AU6"), _
                Array("R8", "AU8"), _
                Array("R10", "AU10"), _
                Array("R12", "AU12"), _
                Array("R14", "AU14"), _
                Array("R16", "AU16"), _
                Array("R18", "AU18"), _
                Array("R20", "AU20"), _
                Array("R22", "AU22"), _
                Array("R24", "AU24"), _
                Array("R26", "AU26"), _
                Array("R28", "AU28"), _
                Array("R30", "AU30"))
        Case 36
            magentaPos = Array( _
                Array("T4", "T32"), _
                Array("V4", "V32"), _
                Array("X4", "X32"), _
                Array("Z4", "Z32"), _
                Array("AB4", "AB32"), _
                Array("AD4", "AD32"), _
                Array("AF4", "AF32"), _
                Array("AH4", "AH32"), _
                Array("AJ4", "AJ32"), _
                Array("AL4", "AL32"), _
                Array("AN4", "AN32"), _
                Array("AP4", "AP32"), _
                Array("AR4", "AR32"), _
                Array("AT4", "AT32"))
        Case 37
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            magentaPos = Array(Array("AB3", "AB32"), Array("AK3", "AK32"), Array("R13", "AU13"), Array("R22", "AU22"))
        Case 38
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            brownPos = Array("S9", "AT9")
        Case 39
            purplePosV = Array("S11", "AT23", "AA4", "AL31")
            purplePosH = Array("AA4", "AL31", "S11", "AT23")
            brownPos = Array("S9", "AT9")
        Case 40
            brownPos = Array("S9", "S26", "AT9", "AT26")
            redPos = Array("S4", "AT4")
        Case 41
            Range("W9:X10").Interior.color = RGB(0, 0, 0)
            Range("W17:X18").Interior.color = RGB(0, 0, 0)
            Range("W25:X26").Interior.color = RGB(0, 0, 0)
            Range("AF9:AG10").Interior.color = RGB(0, 0, 0)
            Range("AF17:AG18").Interior.color = RGB(0, 0, 0)
            Range("AF25:AG26").Interior.color = RGB(0, 0, 0)
            Range("AO9:AP10").Interior.color = RGB(0, 0, 0)
            Range("AO17:AP18").Interior.color = RGB(0, 0, 0)
            Range("AO25:AP26").Interior.color = RGB(0, 0, 0)
            brownPos = Array("S9", "S26", "AT9", "AT26")
        Case 42
            Range("U14:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AJ29").Interior.color = RGB(0, 0, 0)
            Range("AC6:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ6:AJ14").Interior.color = RGB(0, 0, 0)
            Range("AJ14:AR14").Interior.color = RGB(0, 0, 0)
            Range("U21:AC21").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AR21").Interior.color = RGB(0, 0, 0)
            Range("AC21:AC29").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S9", "AT9", "S29", "AT29")
            purplePosH = Array("AF17")
            purplePosV = Array("AF17")
        Case 43
            Range("U14:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AJ29").Interior.color = RGB(0, 0, 0)
            Range("AC6:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ6:AJ14").Interior.color = RGB(0, 0, 0)
            Range("AJ14:AR14").Interior.color = RGB(0, 0, 0)
            Range("U21:AC21").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AR21").Interior.color = RGB(0, 0, 0)
            Range("AC21:AC29").Interior.color = RGB(0, 0, 0)
            magentaPos = Array(Array("U6", "U29"), Array("U29", "AR29"), Array("U6", "AR6"), Array("AR6", "AR29"), Array("Y10", "Y25"), Array("Y25", "AN25"), Array("Y10", "AN10"), Array("AN10", "AN25"))
        Case 44
            Range("U14:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AJ29").Interior.color = RGB(0, 0, 0)
            Range("AC6:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ6:AJ14").Interior.color = RGB(0, 0, 0)
            Range("AJ14:AR14").Interior.color = RGB(0, 0, 0)
            Range("U21:AC21").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AR21").Interior.color = RGB(0, 0, 0)
            Range("AC21:AC29").Interior.color = RGB(0, 0, 0)
            magentaPos = Array(Array("U6", "U29"), Array("U29", "AR29"), Array("U6", "AR6"), Array("AR6", "AR29"), Array("Y10", "Y25"), Array("Y25", "AN25"), Array("Y10", "AN10"), Array("AN10", "AN25"))
            redPos = Array("S4", "AT4")
        Case 45
            Range("U14:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AJ29").Interior.color = RGB(0, 0, 0)
            Range("AC6:AC14").Interior.color = RGB(0, 0, 0)
            Range("AJ6:AJ14").Interior.color = RGB(0, 0, 0)
            Range("AJ14:AR14").Interior.color = RGB(0, 0, 0)
            Range("U21:AC21").Interior.color = RGB(0, 0, 0)
            Range("AJ21:AR21").Interior.color = RGB(0, 0, 0)
            Range("AC21:AC29").Interior.color = RGB(0, 0, 0)
            purplePosH = Array("AF17")
            purplePosV = Array("AF17")
            redPos = Array("S4", "AT4")
            magentaPos = Array(Array("X3", "X32"), Array("AO3", "AO32"), Array("R9", "AU9"), Array("R26", "AU26"))
        Case 46
            Range("Y7:AN7").Interior.color = RGB(0, 0, 0)
            Range("Y11:AN11").Interior.color = RGB(0, 0, 0)
            Range("Y15:AN15").Interior.color = RGB(0, 0, 0)
            Range("Y20:AN20").Interior.color = RGB(0, 0, 0)
            Range("Y24:AN24").Interior.color = RGB(0, 0, 0)
            Range("Y28:AN28").Interior.color = RGB(0, 0, 0)
            purplePosH = Array("AT5", "S9", "AT13", "S17", "AT18", "S22", "AT26", "S30")
        Case 47
            Range("Y7:AN7").Interior.color = RGB(0, 0, 0)
            Range("Y11:AN11").Interior.color = RGB(0, 0, 0)
            Range("Y15:AN15").Interior.color = RGB(0, 0, 0)
            Range("Y20:AN20").Interior.color = RGB(0, 0, 0)
            Range("Y24:AN24").Interior.color = RGB(0, 0, 0)
            Range("Y28:AN28").Interior.color = RGB(0, 0, 0)
            magentaPos = Array(Array("Y7", "Y28"), Array("AN7", "AN28"))
            purplePosH = Array("S17", "AT18")
            purplePosV = Array("U4", "AR4", "AQ31", "V31")
        Case 48
            Range("Y7:AN7").Interior.color = RGB(0, 0, 0)
            Range("Y11:AN11").Interior.color = RGB(0, 0, 0)
            Range("Y15:AN15").Interior.color = RGB(0, 0, 0)
            Range("Y20:AN20").Interior.color = RGB(0, 0, 0)
            Range("Y24:AN24").Interior.color = RGB(0, 0, 0)
            Range("Y28:AN28").Interior.color = RGB(0, 0, 0)
            bluePos = Array("S4", "S8", "S17", "S21", "AT4", "AT8", "AT17", "AT21", "S31", "AT31")
        Case 49
            magentaPos = Array(Array("R16", "AU16"), Array("R17", "AU17"), Array("R9", "AU9"), Array("R10", "AU10"), Array("R23", "AU23"), Array("R24", "AU24"), Array("Z3", "Z32"), Array("Y3", "Y32"), Array("AM3", "AM32"), Array("AN3", "AN32"))
        Case 50
            magentaPos = Array(Array("R16", "AU16"), Array("R17", "AU17"), Array("R9", "AU9"), Array("R10", "AU10"), Array("R23", "AU23"), Array("R24", "AU24"))
            bluePos = Array("S9", "AT9", "S29", "AT29")
        Case 51
            snakePos = Array("W15") ' Pozi?ia de start pentru ?arpe
        Case 52
            snakePos = Array("W15", "W18") ' Pozi?ia de start pentru ?arpe
        Case 53
            bigRedPos = Array("W4")
        Case 54
            invaderPos = Array("W15")
        Case Else
            MsgBox "You finished the game!", vbExclamation
            Exit Sub
    End Select

    ' Call GenerateEnemy to initialize enemies based on the current level
    GenerateEnemy redPos, bluePos, purplePosH, purplePosV, brownPos, magentaPos, purplePosV2, snakePos, bigRedPos, invaderPos
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
    Call StopTimerMagenta
    Call StopTimerPurpleV2  'Opreste timer-ul pentru PurpleV2
    Call StopTimerSpaceInvader

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
Sub GenerateEnemy(redPositions As Variant, bluePositions As Variant, purplePositionsH As Variant, purplePositionsV As Variant, brownPositions As Variant, magentaPositions As Variant, Optional purplePositionsV2 As Variant, Optional snakePositions As Variant, Optional bigRedPositions As Variant, Optional invaderPositions As Variant)
    Dim redCell As redCell
    Dim blueCell As blueCell
    Dim purpleCellH As PurpleCell ' Horizontal purple cell
    Dim purpleCellV As PurpleCell ' Vertical purple cell
    Dim MagentaCell As MagentaCell
    Dim brownCell As brownCell ' Brown cell
    Dim i As Integer
    Dim targetCell1 As Range
    Dim PurpleCellV2 As PurpleCellV2 'Pentru PurpleCellv2
    Dim snakeObject As Snake
    Dim bigRedEnemyObject As BigRedEnemy
    Dim invaderObject As SpaceInvader

    ' Set the initial target cell to the current selection
    Set targetCell1 = ActiveCell

    ' Clear previous collections if they exist
    On Error Resume Next
    Set RedCells = Nothing
    Set BlueCells = Nothing
    Set PurpleCells = Nothing
    Set BrownCells = Nothing
    Set MagentaCells = Nothing ' Initialize brown cell collection
    Set PurpleCellsV2 = Nothing ' Initializare colectie PurpleCellsv2.
        On Error GoTo 0

    Set RedCells = New Collection
    Set BlueCells = New Collection
    Set PurpleCells = New Collection
    Set BrownCells = New Collection
    Set MagentaCells = New Collection
    Set PurpleCellsV2 = New Collection ' Ini?ializare colec?ie PurpleCellv2
    Set Snakes = New Collection
    Set BigRedEnemies = New Collection
    Set SpaceInvaders = New Collection

    ' Check if redPositions array is not empty and initialize red cell objects
    If Not IsEmpty(redPositions) Then
        For i = LBound(redPositions) To UBound(redPositions)

                Set redCell = New redCell
                redCell.Initialize Range(redPositions(i)), IIf(i Mod 2 = 0, True, False) ' Example: alternate movement directions
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
            Set MagentaCell = New MagentaCell
            MagentaCell.Initialize magentaPositions(i) ' Initialize magenta cell with string
            MagentaCells.Add MagentaCell
        Next i
    End If

    'Ini?ializeaza PurpleCellv2 objects:
      If Not IsEmpty(purplePositionsV2) Then
        For i = LBound(purplePositionsV2) To UBound(purplePositionsV2)
            Dim isVertical As Boolean
            isVertical = Not (IsEmpty(purplePositionsH) Or IsEmpty(purplePositionsV)) 'Verificam daca exista celule PurpleCell normale. Daca exista, initializam PurpleCellv2 cu directii opuse primei celule PurpleCell create.

                Set PurpleCellV2 = New PurpleCellV2 ' Folose?te PurpleCellv2
                PurpleCellV2.Initialize Range(purplePositionsV2(i)), isVertical
                PurpleCellsV2.Add PurpleCellV2  'Adauga în colec?ia PurpleCellsV2
                With Range(purplePositionsV2(i)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
                    .Interior.color = RGB(128, 0, 128)
                End With
        Next i
    End If
    
    If Not IsEmpty(snakePositions) Then
        For i = LBound(snakePositions) To UBound(snakePositions)
            Set snakeObject = New Snake
            snakeObject.Initialize Range(snakePositions(i)), IIf(i Mod 2 = 0, False, True)
            Snakes.Add snakeObject
        Next i
    End If
    
    If Not IsEmpty(bigRedPositions) Then
        For i = LBound(bigRedPositions) To UBound(bigRedPositions)
            Set bigRedEnemyObject = New BigRedEnemy
            bigRedEnemyObject.Initialize Range(bigRedPositions(i)), IIf(i Mod 2 = 0, False, True)
            BigRedEnemies.Add bigRedEnemyObject
        Next i
    End If
    If Not IsEmpty(invaderPositions) Then
        For i = LBound(invaderPositions) To UBound(invaderPositions)
            Set invaderObject = New SpaceInvader
            invaderObject.Initialize Range(invaderPositions(i))
            SpaceInvaders.Add invaderObject
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
    TimerIDMagenta = SetTimer(0, 0, MoveIntervalMagenta, AddressOf TimerEventMagenta)
End Sub

' Stop timer for magenta cells
Public Sub StopTimerMagenta()
    If TimerIDMagenta <> 0 Then
        KillTimer 0, TimerIDMagenta
        TimerIDMagenta = 0
    End If
End Sub

' Handler for magenta cell movements
Sub TimerEventMagenta()
    On Error Resume Next
    Call MoveMagentaCells
End Sub
Sub MoveMagentaCells()
    Dim MagentaCell As MagentaCell
    Dim i As Integer

    For i = 1 To MagentaCells.count
        Set MagentaCell = MagentaCells(i)
        MagentaCell.Move
    Next i
End Sub

'Adauga func?iile pentru timerul PurpleCellv2:
Public Sub StartTimerPurpleV2()
    If TimerIDPurpleV2 <> 0 Then KillTimer 0, TimerIDPurpleV2
    TimerIDPurpleV2 = SetTimer(0, 0, MoveIntervalPurpleV2, AddressOf TimerEventPurpleV2)
End Sub

Public Sub StopTimerPurpleV2()
    If TimerIDPurpleV2 <> 0 Then KillTimer 0, TimerIDPurpleV2
End Sub

Sub TimerEventPurpleV2()
    On Error Resume Next
    Call MovePurpleCellsV2
End Sub

Sub MovePurpleCellsV2()
    Dim PurpleCellV2 As PurpleCellV2
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To PurpleCellsV2.count
        Set PurpleCellV2 = PurpleCellsV2(i)
        PurpleCellV2.Move
        If PurpleCellV2.PurpleCell.Address = ActiveCell.Address Then anyTargetReached = True
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

Public Sub StartTimerSnake()
    If TimerIDSnake <> 0 Then KillTimer 0, TimerIDSnake
    MoveIntervalSnake = 150 ' Seta?i intervalul de mi?care dorit
    TimerIDSnake = SetTimer(0, 0, MoveIntervalSnake, AddressOf TimerEventSnake)
End Sub

Public Sub StopTimerSnake()
    If TimerIDSnake <> 0 Then KillTimer 0, TimerIDSnake
End Sub

Sub TimerEventSnake()
    On Error Resume Next
    Call MoveSnakes
End Sub

Sub MoveSnakes()
    Dim snakeObject As Snake
    Dim anyTargetReached As Boolean
    Dim i As Integer

    anyTargetReached = False

    For i = 1 To Snakes.count
        Set snakeObject = Snakes(i)
        snakeObject.Move ActiveCell ' Muta?i ?arpele în direc?ia celulei active
        If snakeObject.SnakeCell.Address = ActiveCell.Address Then
            anyTargetReached = True
        End If
    Next i

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

Public Sub StartTimerBigRed()
    If TimerIDBigRed <> 0 Then KillTimer 0, TimerIDBigRed
    TimerIDBigRed = SetTimer(0, 0, MoveIntervalBigRed, AddressOf TimerEventBigRed)
End Sub

Public Sub StopTimerBigRed()
    If TimerIDBigRed <> 0 Then KillTimer 0, TimerIDBigRed
End Sub

Sub TimerEventBigRed()
    On Error Resume Next
    Call MoveBigRedEnemies
End Sub

Sub MoveBigRedEnemies()
    Dim bigRed As BigRedEnemy
    For Each bigRed In BigRedEnemies
        bigRed.Move ActiveCell
        ' Pute?i adauga aici logica de coliziune daca este necesar
    Next bigRed
End Sub
' Timer handling and movement logic for Space Invaders
Public Sub StartTimerSpaceInvader()
    If TimerIDSpaceInvader <> 0 Then KillTimer 0, TimerIDSpaceInvader
    TimerIDSpaceInvader = SetTimer(0, 0, MoveIntervalSpaceInvader, AddressOf TimerEventSpaceInvader)
End Sub

Public Sub StopTimerSpaceInvader()
    If TimerIDSpaceInvader <> 0 Then KillTimer 0, TimerIDSpaceInvader
End Sub

Sub TimerEventSpaceInvader()
    On Error Resume Next
    Call MoveSpaceInvaders
End Sub

Sub MoveSpaceInvaders()
    Dim invader As SpaceInvader
    Dim anyTargetReached As Boolean
    anyTargetReached = False

    For Each invader In SpaceInvaders
        invader.Move
        ' Verifica coliziunea cu player-ul
        If Not Intersect(ActiveCell, invader.EnemyArea) Is Nothing Then
            anyTargetReached = True
            Exit For ' Ie?i din bucla daca s-a detectat coliziunea
        End If
    Next invader

    If anyTargetReached And go = True Then
        go = False
        StopAllTimers
        MsgBox "Game Over"
        ResetGame
        StartGame
    End If
End Sub

' Stop all timers
Public Sub StopAllTimers()
    StopTimerRed
    StopTimerBlue
    StopTimerPurple
    StopTimerBrown
    StopTimerMagenta
    StopTimerPurpleV2 'Opreste si timerul pentru PurpleCellV2
    StopTimerSnake
    StopTimerBigRed
    StopTimerSpaceInvader
    Call StopTimerSelection
End Sub

Sub levelup()
level = level + 1
End Sub

