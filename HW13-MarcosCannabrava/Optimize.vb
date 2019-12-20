Module Optimize

    Public Sub SpeedUp()

        With Excel.Application
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
        End With

    End Sub


    Sub ResetConfig()

        With Excel.Application
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = True
        End With

    End Sub


End Module
