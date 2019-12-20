Module Synching

    'global variables -------------------------------------------
    Public CurrentDate As Date = "1/1/1"
    Public LastTransactionDate As Date = "1/1/1"
    Public HTstartDate As Date = "1/1/1"

    Public TeamID As String = "03"
    Public TeamPortfolioTableName = "PortfolioTeam" & TeamID
    Public TeamConfirmationTicketTableName = "ConfirmationTicketTeam" & TeamID
    Public excludeIPforTesting As Boolean = False

    Public WithEvents MainTimer As Timer
    Public WithEvents ScreenTimer As Timer
    Public secondsLeft As Integer
    Public autopilot As Boolean = False
    Public mainTimerEngaged As Boolean = False

    Public Volatilities As Double() = {0.242, 0.229, 0.132, 0.294, 0.196, 0.151, 0.206, 0.162, 0.13, 0.112, 0.132, 0.104}

    Public suggestedSymbol As String
    Public suggestedAction As String
    Public suggestedSignedQty As Double

    Public startCounting As DateTime

    '-------------------------------------------------------------

    Public Sub InitializeDBSession()
        ClearDS()
        Globals.Sheet1.ClearListObjects()
        Globals.Sheet2.ClearListObjects()
        Globals.Sheet3.ClearListObjects()
        Globals.Sheet4.ClearListObjects()
        Globals.Sheet5.ClearListObjects()

        CurrentDate = GetCurrentDate()
        HTstartDate = GetHTstartDate()

        Globals.Sheet1.InitializeDashboard()
    End Sub

    Public Sub DailyRoutine()
        Dim targetDate As Date = CurrentDate
        If (CurrentDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = CurrentDate.AddDays(-1)
        End If
        If (CurrentDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = CurrentDate.AddDays(-2)
        End If

        RunQueryAndSaveResultsInDS("select * from StockMarket where date = '" & targetDate.ToShortDateString() & "'", _
                                   "StockMarketForOneDayTbl")
        RunQueryAndSaveResultsInDS("select * from OptionMarket where date = '" & targetDate.ToShortDateString() & "'", _
                                   "OptionMarketForOneDayTbl")
        Globals.Sheet1.ResetTransactionData()
        Globals.Sheet1.DisplayTransactionData()
        Globals.Sheet1.ReCalcFinancialMetrics()
        Globals.Sheet1.DisplayFinancialMetrics()
        Globals.Sheet5.AddRowToTETable(CurrentDate)

        'If CurrentDate = "07/20/2013" Or "10/19/2013" Then
        '    ExerciseOptions()                                       'WRITE THIS !!!!!!!!!!
        'End If

    End Sub

    Public Sub SetUpTimers()
        MainTimer = New Timer
        MainTimer.Interval = 3000
        ScreenTimer = New Timer
        ScreenTimer.Interval = 1000
        'SetSeconds(60)
    End Sub

    Public Sub SetSeconds(sec As Integer)
        Try
            secondsLeft = sec
            If Globals.ThisWorkbook.ActiveSheet.name = "Dashboard" Then
                Globals.Sheet1.SecondsLeftCell.Value = sec
                Globals.Sheet1.SecondsLeftCell.Font.Size = 60
            End If
        Catch ex As Exception
            '
        End Try
    End Sub

    Private Sub MainTimer_tick() Handles MainTimer.Tick
        Dim tempNewDate As Date
        tempNewDate = GetCurrentDate()
        If tempNewDate.Date <> CurrentDate.Date Then
            CurrentDate = tempNewDate
            startCounting = DateTime.Now()

            If ScreenTimer.Enabled Then '''''''''''''
                SetSeconds(60) ''''''''''''
            Else '''''''''''
                ScreenTimer.Start() ''''''''''''''''
                SetSeconds(60) ''''''''''''
            End If ''''''''''''''
            DailyRoutine()
            If autopilot Then
                AutoPilotRoutine()
            End If
        End If
    End Sub

    Private Sub ScreenTimer_Tick() Handles ScreenTimer.Tick
        Try
            Dim ts As TimeSpan = DateTime.Now() - startCounting
            If Globals.ThisWorkbook.ActiveSheet.Name = "Dashboard" And
                (60 - ts.Seconds) > -0 Then
                Globals.Sheet1.SecondsLeftCell.Value = 60 - ts.Seconds
            End If
        Catch ex As Exception
            '
        End Try
        If secondsLeft > 0 Then  '''''''''''''
            SetSeconds(secondsLeft - 1) '''''''''''''''
        End If ''''''''''''''''''
    End Sub

    Public Sub AutoPilotRoutine()
        CalcHedge()
        'Globals.Sheet1.Application.ScreenUpdating = False
        'For i = 1 To 12
        '    Globals.Sheet1.ExecuteHedge(i)
        'Next
        'Globals.Sheet1.Application.ScreenUpdating = True
    End Sub

End Module
