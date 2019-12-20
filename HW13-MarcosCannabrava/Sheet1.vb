
Public Class Sheet1


    Private Sub Sheet1_Startup() Handles Me.Startup
        Me.Activate()
        Application.ActiveWindow.DisplayHeadings = False
        Application.ActiveCell.Font.Size = 9
        TransactionQueueLst.AutoSetDataBoundColumnHeaders = True
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    'VARIABLES -----------------------------
    Public tDate As Date
    Public tType As String
    Public tSymbol As String
    Public tQty As Integer
    Public tPrice As Double
    Public totTcost As Double
    Public totTvalue As Double
    Public percTcost As Double
    Public interestSLT As Double
    Public cAccountAT As Double
    Public marginsAT As Double
    Public cAccount As Double

    Public maxMargins As Double
    Public TPVatStart As Double
    Public taTPV As Double
    Public iRate As Double
    Public TPV As Double
    Public TEweighted As Double
    Public TEpercent As Double
    Public margins As Double


    Public Sub ClearListObjects()
        TransactionQueueLst.DataBodyRange.Clear()
        TransactionQueueLst.DataBodyRange.Font.Size = 9
    End Sub

    Public Sub InitializeDashboard()
        ' ResetTransactionData()

        ' DisplayTransactionData()
        ClearListObjects()
        'DisplayFinancialMetrics() 'CHECK IT
        FillCBoxes()
        DownloadTransactionQueueForTeam()
        RunQueryAndSaveResultsInDS("select * from TransactionCost", "TransactionCostTbl")
        RunQueryAndSaveResultsInDS("select * from InitialPosition", "InitialPositionTbl")
        Globals.Sheet5.InitializeCharts()
        LastTransactionDate = GetLastTransactionDate()


        InitializeSmartHedger()
        DailyRoutine()
        ResetFinancialMetrics()

    End Sub

    Public Sub ResetTransactionData()
        tDate = CurrentDate
        tType = ""
        tSymbol = ""
        tQty = 0
        tPrice = 0
        totTcost = 0
        totTvalue = 0
        percTcost = 0
        interestSLT = 0
        cAccountAT = 0
        marginsAT = 0
    End Sub

    Public Sub ResetFinancialMetrics()
        cAccount = GetPositionFromDB("CAccount")
        maxMargins = GetMaxMargins()
        iRate = GetIRate()
        TPVatStart = Globals.Sheet3.CalcTPVatStart()
    End Sub

    Public Sub ReCalcFinancialMetrics()
        taTPV = Globals.Sheet3.calcTaTPV(CurrentDate)
        TPV = Globals.Sheet3.CalcTPV(CurrentDate)
        TEweighted = Globals.Sheet3.CalcWeightedTE(CurrentDate)

    End Sub

    Public Sub DisplayTransactionData()
        CurrentDateCell.Value = CurrentDate.ToLongDateString()
        TransactionDateCell.Value = tDate
        TypeCell.Value = tType
        SymbolCell.Value = tSymbol
        QtyCell.Value = tQty
        PriceCell.Value = tPrice
        PriceCell.NumberFormat = "$##,###.##"
        TotalTCostCell.Value = totTcost
        TotalTCostCell.NumberFormat = "$##,###,###.##"
        TotalTValueCell.Value = totTvalue
        TotalTValueCell.NumberFormat = "$##,###,###.##"


        InterestSLTCell.Value = interestSLT
        InterestSLTCell.NumberFormat = "$##,###,###.##"
        cAcctATCell.Value = cAccountAT
        cAcctATCell.NumberFormat = "$##,###,###.##"
        Margin_ATCell.Value = marginsAT
        Margin_ATCell.NumberFormat = "$##,###,###.##"

    End Sub

    Public Sub DisplayFinancialMetrics()
        cAcctCell.Value = cAccount
        cAcctCell.NumberFormat = "$##,###,###.##"

        MaxMarginsCell.Value = maxMargins
        MaxMarginsCell.NumberFormat = "$##,###,###.##"
        TPV_at_StartCell.Value = TPVatStart
        TPV_at_StartCell.NumberFormat = "$##,###,###.##"
        TaTPVCell.Value = taTPV
        TaTPVCell.NumberFormat = "$##,###,###.##"
        TPVCell.Value = TPV
        TPVCell.NumberFormat = "$##,###,###.##"
        TE__weightedCell.Value = TEweighted
        TE__weightedCell.NumberFormat = "$##,###,###.##"
        TECell.Value = TEweighted / taTPV
        TECell.NumberFormat = "##.##%"
        MarginsCell.Value = margins
        MarginsCell.NumberFormat = "$##,###,###.##"
        _30MarginsCell.Value = margins * 0.3
        _30MarginsCell.NumberFormat = "$##,###,###.##"

    End Sub

    Public Sub FillCBoxes()
        Dim q As String

        'filling TickerCBox
        q = "select ticker from StockTicker order by ticker"
        RunQueryAndSaveResultsInDS(q, "TickerTbl")
        TickerCBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("TickerTbl").Rows
            TickerCBox.Items.Add(dr("Ticker").ToString().Trim())
        Next
        TickerCBox.Text = "Stock"

        'filling SymbolCBox
        q = "select symbol from OptionSymbol order by symbol"
        RunQueryAndSaveResultsInDS(q, "SymbolTbl")
        SymbolCBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("SymbolTbl").Rows
            SymbolCBox.Items.Add(dr("Symbol").ToString().Trim())
        Next
        SymbolCBox.Text = "Option"

        'filling StockTTypeCBox
        StockTTypeCBox.Items.Clear()
        For Each t As String In {"Buy", "Sell", "SellShort", "CashDiv"}
            StockTTypeCBox.Items.Add(t)
        Next
        StockTTypeCBox.Text = "Select"

        'filling OptionTTypeCBox
        OptionTTypeCBox.Items.Clear()
        For Each t As String In {"Buy", "Sell", "SellShort", "X-Call", "X-Put"}
            OptionTTypeCBox.Items.Add(t)
        Next
        OptionTTypeCBox.Text = "Select"
    End Sub

    Private Sub DownloadTransactionQueueForTeam()
        Dim m = String.Format("select * from TransactionQueue where teamID = '{0}' order by rowID desc", TeamID)
        RunQueryAndSaveResultsInDS(m, "TransactionQueueTbl")
        TransactionQueueLst.DataSource = myDataSet.Tables("TransactionQueueTbl")
    End Sub


    'PREVIEW STOCK MARKET TRANSACTION -----------------------------------------------------------
    Private Sub StockTradePreviewBtn_Click(sender As Object, e As EventArgs) Handles StockTradePreviewBtn.Click
        tType = StockTTypeCBox.SelectedItem.ToString().Trim()
        tQty = Integer.Parse(StockQtyTBox.Text)
        tSymbol = TickerCBox.SelectedItem

        tDate = CurrentDate

        CalcTransactionPreview(tDate)
        DisplayTransactionData()

    End Sub

    Private Sub CalcTransactionPreview(targetDate As Date)

        Select Case tType
            Case "Buy"
                tPrice = FindPrice(tSymbol, "Ask", targetDate)
            Case "Sell"
                tPrice = FindPrice(tSymbol, "Bid", targetDate)
            Case "SellShort"
                tPrice = FindPrice(tSymbol, "Bid", targetDate)
            Case "CashDiv"
                tPrice = FindDividend(tSymbol, targetDate)
            Case "X-Call"
                tPrice = GetStrike(tSymbol)
            Case "X-Put"
                tPrice = GetStrike(tSymbol)
        End Select

        If IsAStock(tSymbol) Then
            totTcost = GetTransactionCostCoefficient("Stock", tType) * tQty * tPrice
        Else
            totTcost = GetTransactionCostCoefficient("Option", tType) * tQty * tPrice
        End If


        totTvalue = ComputeTotTValue()
        interestSLT = Globals.Sheet3.CalcInterestSLT(CurrentDate)
        cAccountAT = cAccount + interestSLT + totTvalue
        marginsAT = margins + Globals.Sheet3.CalcEffectOfTransactionOnMargin(tType, tSymbol, tQty)

    End Sub

    Private Function ComputeTotTValue() As Double
        Select Case tType
            Case "Buy"
                Return -(tPrice * tQty) - totTcost
            Case "Sell", "SellShort"
                Return (tPrice * tQty) - totTcost
            Case "CashDiv"
                Return (tPrice * tQty) - totTcost
            Case "X-Call"
                Return -(tPrice * tQty) - totTcost
            Case "X-Put"
                Return (tPrice * tQty) - totTcost
        End Select
        Return 0
    End Function

    'EXECUTE STOCK TRANSACTION -------------------------------------------------------------------
    Private Sub StockTradeBtn_Click(sender As Object, e As EventArgs) Handles StockTradeBtn.Click
        StockTradePreviewBtn_Click(Nothing, Nothing)
        ExecuteTrade()
        DisplayTransactionData()

        DisplayFinancialMetrics()
        DownloadTransactionQueueForTeam()
    End Sub

    Private Sub ExecuteTrade()
        If IsTradeValid() = False Then
            MessageBox.Show("Invalid Transaction.")
            Exit Sub
        End If
        Dim q = String.Format("INSERT INTO TransactionQueue (Date, TeamID, Symbol, Type, Qty, Price, Cost, TotValue, " & _
                              "InterestSinceLastTransaction, CashPositionAfterTransaction, TotMargin) VALUES " & _
                              "('{0}', {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10})", _
                              CurrentDate.ToShortDateString(), TeamID, tSymbol, tType, tQty, tPrice, totTcost, totTvalue, interestSLT, cAccountAT, marginsAT)
        ExecuteNonQuery(q)
        LastTransactionDate = CurrentDate
        cAccount = cAccountAT
        Globals.Sheet3.UpdatePortfolio(tType, tSymbol, tQty, (totTvalue + interestSLT))
        margins = marginsAT
        ReCalcFinancialMetrics()
    End Sub


    'PREVIEW OPTION MARKET TRANSACTION -----------------------------------------------------------
    Private Sub OptionTradePreviewBtn_Click(sender As Object, e As EventArgs) Handles OptionTradePreviewBtn.Click
        tType = OptionTTypeCBox.SelectedItem.ToString().Trim()
        tQty = Integer.Parse(OptionQtyTBox.Text)
        tSymbol = SymbolCBox.SelectedItem

        tDate = CurrentDate

        CalcTransactionPreview(tDate)
        DisplayTransactionData()

    End Sub

    'EXECUTE OPTION MARKET TRANSACTION -----------------------------------------------------------
    Private Sub OptionTradeBtn_Click(sender As Object, e As EventArgs) Handles OptionTradeBtn.Click
        OptionTradePreviewBtn_Click(Nothing, Nothing)
        ExecuteTrade()
        DisplayTransactionData()

        DisplayFinancialMetrics()
        DownloadTransactionQueueForTeam()
    End Sub

    Public Sub ExecuteHedge(row As Integer)
        tType = RecommendationRange.Cells(row, 1).value.ToString().Trim()
        If tType = "Hold" Or tType = "-" Then
            Exit Sub
        End If
        tQty = Math.Abs(QtyRange.Cells(row, 1).value)
        tSymbol = SymbolRange.Cells(row, 1).Value.ToString().Trim()
        tDate = CurrentDate
        CalcTransactionPreview(tDate)
        DisplayTransactionData()
        ExecuteTrade()
        DownloadTransactionQueueForTeam()
        DailyRoutine()
        CalcHedge()
    End Sub

    Private Sub TradeItBtn_Click(sender As Object, e As EventArgs) Handles TradeItBtn.Click
        ExecuteHedge(1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ExecuteHedge(2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ExecuteHedge(3)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ExecuteHedge(12)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ExecuteHedge(4)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ExecuteHedge(5)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExecuteHedge(6)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ExecuteHedge(7)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ExecuteHedge(8)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ExecuteHedge(9)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ExecuteHedge(10)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ExecuteHedge(11)
    End Sub

    'Public Sub ExerciseOptions()

    '    Dim q As String = "Select * from " & TeamPortfolioTableName
    '    RunQueryAndSaveResultsInDS(q, TeamPortfolioTableName)

    '    If CurrentDate = "07/20/2013" Then

    '    End If

    '    tType = OptionTTypeCBox.SelectedItem.ToString().Trim()
    '    tQty = Integer.Parse(OptionQtyTBox.Text)
    '    tSymbol = SymbolCBox.SelectedItem
    '    tDate = CurrentDate

    '    CalcTransactionPreview(tDate)
    '    ExecuteTrade()
    '    DisplayTransactionData()
    '    DisplayFinancialMetrics()
    '    DownloadTransactionQueueForTeam()
    'End Sub

End Class
