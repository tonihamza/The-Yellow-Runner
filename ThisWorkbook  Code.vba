Private Sub Workbook_Open()
    ' Set up key bindings when the workbook is opened
    Application.OnKey "{UP}", "MoveUp"
    Application.OnKey "{DOWN}", "MoveDown"
    Application.OnKey "{LEFT}", "MoveLeft"
    Application.OnKey "{RIGHT}", "MoveRight"
    Application.OnKey " ", "StartLevel"
    nivel = 1 ' Initialize the level
    Call StartGame
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Remove key bindings before closing the workbook
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey " "
    
    ' Ensure timers are stopped when closing
    StopTimerSelection
    StopTimerRed
    StopTimerBlue
End Sub

