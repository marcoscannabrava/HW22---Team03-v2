Imports Microsoft.Office.Tools.Ribbon

Public Class SpartanTraderRibbon

    Private Sub SpartanTraderRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Me.RibbonUI.ActivateTabMso("TabAddIns")
        InitializeWorkstation()
        DoAtStart()
    End Sub

    Private Sub InitializeWorkstation()
        ' things that do not change when we switch database
        SetUpADOcomponents()
        Globals.ThisWorkbook.Application.ActiveWindow.DisplayFormulas = False
    End Sub

    Private Sub DoAtStart()
        'AlphaTBtn_Click(Nothing, Nothing)
        'Globals.Sheet1.InitializeDashboard()
        DashboardBtn_Click(Nothing, Nothing)
    End Sub

    Private Sub AlphaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaTBtn.Click
        If mainTimerEngaged Then
            MainTimerToggleBtn_Click(Nothing, Nothing)
        End If
        DisconnectFromDB()
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = False
        ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;Initial Catalog=HedgeTournamentAlpha;Integrated Security=true")
        If NoData() Then
            MessageBox.Show("There is no data in this Database")
            Exit Sub
        End If
        AlphaTBtn.Checked = True
        InitializeDBSession()
    End Sub

    Private Sub BetaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaTBtn.Click
        If mainTimerEngaged Then
            MainTimerToggleBtn_Click(Nothing, Nothing)
        End If
        DisconnectFromDB()
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = False
        ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;Initial Catalog=HedgeTournamentBeta;Integrated Security=true")
        If NoData() Then
            MessageBox.Show("There is no data in this Database")
            Exit Sub
        End If
        BetaTBtn.Checked = True
        InitializeDBSession()
    End Sub

    Private Sub QuitBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitBtn.Click
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.Quit()
    End Sub

    Private Sub StocksBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StocksBtn.Click
        Globals.Sheet2.Activate()
        Dim q = String.Format("select * from StockMarket where date = '{0}'", CurrentDate.ToShortDateString)
        RunQueryAndSaveResultsInDS(q, "StockTbl")
        Globals.Sheet2.StocksLst.DataSource = myDataSet.Tables("StockTbl")
    End Sub

    Private Sub OptionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionsBtn.Click
        Globals.Sheet2.Activate()
        Dim q = String.Format("select * from OptionMarket where date = '{0}'", CurrentDate.ToShortDateString)
        RunQueryAndSaveResultsInDS(q, "OptionsTbl")
        Globals.Sheet2.OptionsLst.DataSource = myDataSet.Tables("OptionsTbl")
    End Sub

    Private Sub IndexBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles IndexBtn.Click
        Globals.Sheet2.Activate()
        Dim q = "select * from StockIndex"
        RunQueryAndSaveResultsInDS(q, "IndexTbl")
        Globals.Sheet2.IndexLst.DataSource = myDataSet.Tables("IndexTbl")
    End Sub

    Private Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click
        Globals.Sheet1.Activate()
    End Sub

    Private Sub InitialPositionBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionBtn.Click
        Globals.Sheet3.Activate()
        Dim q = "select * from InitialPosition"
        RunQueryAndSaveResultsInDS(q, "InitialPositionTbl")
        Globals.Sheet3.InitialPositionLst.DataSource = myDataSet.Tables("InitialPositionTbl")
    End Sub

    Private Sub PortfolioBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles PortfolioBtn.Click
        Globals.Sheet3.Activate()
        Dim q = "select * from " & TeamPortfolioTableName
        RunQueryAndSaveResultsInDS(q, "TeamPortfolioTbl")
        Globals.Sheet3.TeamPortfolioLst.DataSource = myDataSet.Tables("TeamPortfolioTbl")
    End Sub

    Private Sub ConfirmTicketBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfirmTicketBtn.Click
        Globals.Sheet3.Activate()
        Dim q = "select * from " & TeamConfirmationTicketTableName
        RunQueryAndSaveResultsInDS(q, "TeamConfirmationTicketTbl")
        Globals.Sheet3.ConfirmationTicketsLst.DataSource = myDataSet.Tables("TeamConfirmationTicketTbl")
    End Sub

    Private Sub EnvironmentBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles EnvironmentBtn.Click
        Globals.Sheet4.Activate()
        Dim q = "select * from EnvironmentVariable"
        RunQueryAndSaveResultsInDS(q, "EnvironmentVariableTbl")
        Globals.Sheet4.EnvironmentLst.DataSource = myDataSet.Tables("EnvironmentVariableTbl")
    End Sub

    Private Sub TCostsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TCostsBtn.Click
        Globals.Sheet4.Activate()
        Dim q = "select * from TransactionCost"
        RunQueryAndSaveResultsInDS(q, "TransactionCostTbl")
        Globals.Sheet4.TCostLst.DataSource = myDataSet.Tables("TransactionCostTbl")
    End Sub

    Private Sub PFResetBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles PFResetBtn.Click
        Globals.Sheet3.Activate()
        Globals.Sheet3.ResetPortfolio()
        PortfolioBtn_Click(Nothing, Nothing)

    End Sub

    Private Sub PFUpdateBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles PFUpdateBtn.Click
        Globals.Sheet3.Activate()
        Globals.Sheet3.UploadPortfolioToDB()
    End Sub

    Private Sub DateOverrideBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DateOverrideBtn.Click
        If mainTimerEngaged Then
            MainTimerToggleBtn_click(Nothing, Nothing)
        End If

        If IsDate(OverrideBox.Text) Then
            Dim oldCurrentDate As Date = CurrentDate
            CurrentDate = Date.Parse(OverrideBox.Text)
            If LastTransactionDate.Date > CurrentDate Then
                LastTransactionDate = CurrentDate
            End If
            If CurrentDate.Date < oldCurrentDate.Date Then
                Globals.Sheet5.ClearTEdataTbl()
                Globals.Sheet1.Chart_1.SetSourceData(Globals.Sheet5.TEtoChartLST.Range)
            End If
            DailyRoutine()
        End If
    End Sub

    Private Sub ProjectPFBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ProjectPFBtn.Click
        Globals.Sheet5.ProjectPortfolio()
    End Sub

    Private Sub StockOptionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StockOptionsBtn.Click
        Globals.Sheet5.Activate()
    End Sub

    Private Sub ToggleButton2_Click(sender As Object, e As RibbonControlEventArgs) Handles AutoPilotTglBtn.Click
        If mainTimerEngaged Then
            AutoPilotTglBtn.Checked = Not autopilot
            autopilot = Not autopilot
        End If
    End Sub

    Private Sub MainTimerToggleBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles MainTimerToggleBtn.Click
        MainTimerToggleBtn.Checked = Not mainTimerEngaged
        mainTimerEngaged = Not MainTimerEngaged
        autopilot = False
        AutoPilotTglBtn.Checked = False
        If mainTimerEngaged Then
            SetUpTimers()
            MainTimer.Start()
            startCounting = DateTime.Now
            ScreenTimer.Start()
        Else
            MainTimer.Stop()
            ScreenTimer.Stop()
            'SetSeconds(60)
        End If
    End Sub

    Private Sub CalcHedgeBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles CalcHedgeBtn.Click
        DailyRoutine()
        AutoPilotRoutine()
    End Sub
End Class
