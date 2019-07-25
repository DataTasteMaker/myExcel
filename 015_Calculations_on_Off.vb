Sub TurnCalcsOff()
        With Application
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .DisplayAlerts = False
            .ScreenUpdating = False
        End With
End Sub

Sub TurnCalcsOn()
        With Application
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            .DisplayAlerts = True
            .ScreenUpdating = True
        End With
End Sub
