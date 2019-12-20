Module TradingAlgorithm

    Public Sub CalcHedge()
        SpeedUp()
        Dim q As String
        Dim familyTicker As String
        Globals.Sheet1.Range("L4:Q15").Value = "-"
        For i = 1 To 12
            familyTicker = Globals.Sheet1.UnderlierRange.Cells(i, 1).Value.ToString().Trim()
            Globals.Sheet1.PortfolioDeltaRange.Cells(i, 1) = CalcFamilyDelta(familyTicker)
        Next
        If HedgingToday(CurrentDate) Then
            q = "Select * from " & TeamPortfolioTableName
            RunQueryAndSaveResultsInDS(q, TeamPortfolioTableName)
            For i = 1 To 12
                familyTicker = Globals.Sheet1.UnderlierRange.Cells(i, 1).value.ToString().Trim()
                suggestedAction = "Hold"
                suggestedSymbol = "-"
                suggestedSignedQty = 0
                If NeedToHedge(familyTicker) Then
                    SuggestHedge(familyTicker)
                End If
                Globals.Sheet1.RecommendationRange.Cells(i, 1) = suggestedAction
                Globals.Sheet1.SymbolRange.Cells(i, 1) = suggestedSymbol
                Globals.Sheet1.QtyRange.Cells(i, 1) = suggestedSignedQty
            Next
        End If
        ResetConfig()
    End Sub

    Public Sub InitializeSmartHedger()
        Globals.Sheet1.Range("I4:Q15").Value = ""
        DisplayUnderliers()
        DisplayVols()
    End Sub

    Public Sub DisplayUnderliers()
        For i = 1 To 12
            Globals.Sheet1.UnderlierRange.Cells(i, 1) = myDataSet.Tables("TickerTbl").Rows(i - 1)("ticker")
        Next
    End Sub

    Public Sub DisplayVols()
        For i = 1 To 12
            Globals.Sheet1.VolatilityRange.Cells(i, 1) = Volatilities(i - 1)
        Next
    End Sub

    Private Function HedgingToday(d As Date) As Boolean 'PUT HERE THE DAYS WE WANT TO HEDGE
        If d.DayOfWeek = DayOfWeek.Saturday Or
            d.DayOfWeek = DayOfWeek.Sunday Then
            Return False
        End If
        Return True
    End Function

    Private Function NeedToHedge(familyTkr As String) As Boolean ' I CAN CHANGE THE 50000 VALUE
        If Math.Abs(GetFamilyPortfolioDelta(familyTkr)) < 50000 Then
            Return False
        End If
        Return True
    End Function

    Private Function GetFamilyPortfolioDelta(tkr As String) As Double
        For i = 1 To 12
            If Globals.Sheet1.UnderlierRange.Cells(i, 1).Value.trim() = tkr Then
                Return Globals.Sheet1.PortfolioDeltaRange.Cells(i, 1).Value
            End If
        Next
        Return 0
    End Function

    Public Sub SuggestHedge(familyTicker As String)
        'FIRST DRAFT OF TRADING ALGORITHM, WE HAVE TO  IMPROVE
        If GetFamilyPortfolioDelta(familyTicker) > 0 Then
            'Those will be tried in that order, the first that fits will be suggested. Free to change the order
            If CheckSellStock(familyTicker) Then
                Return
            End If
            If CheckSellCall(familyTicker) Then
                Return
            End If
            If CheckBuyBackPut(familyTicker) Then
                Return
            End If
            If CheckBuyPut(familyTicker) Then
                Return
            End If
            If CheckSellShortCall(familyTicker) Then
                Return
            End If
            If CheckSellShortStock(familyTicker) Then
                Return
            End If
            
        Else 'family delta < 0
            If CheckSellPut(familyTicker) Then
                Return
            End If
            If CheckBuyBackCall(familyTicker) Then
                Return
            End If
            If CheckBackStock(familyTicker) Then
                Return
            End If
            If CheckBuyCall(familyTicker) Then
                Return
            End If
            If CheckBuyStock(familyTicker) Then
                Return
            End If
            If CheckSellShortPut(familyTicker) Then
                Return
            End If
        End If
    End Sub

    Public Function CalcSignedQtyNeededToHedge(familyToHedge As String, symbolToUse As String) As Integer
        Dim delta As Double
        Dim deltaTarget As Double = 0
        Dim portfolioDelta = GetFamilyPortfolioDelta(familyToHedge)
        delta = CalcDelta(symbolToUse, CurrentDate)
        If Math.Abs(delta) < 0.05 Then ' Arbitrary, I can change
            Return 0
        End If
        Return Math.Round((deltaTarget - portfolioDelta) / delta)
    End Function

    Public Function MaxShortBuyBack(sym As String) As Double
        Dim q As Double = 0
        Dim availableCash As Double = 0
        If (Globals.Sheet1.cAccount - Globals.Sheet1.margins * 0.3) > 0 Then
            availableCash = ((Globals.Sheet1.cAccount - Globals.Sheet1.margins * 0.3) / 0.7)
            availableCash = availableCash * 0.9 '10 % is a cushion to pay for t costs and safety
            q = availableCash / FindPrice(sym, "Ask", CurrentDate)
            Return Math.Truncate(q)
        Else
            Return 0
        End If
    End Function

    Public Function AvailableCashIsLow() As Boolean
        If ((Globals.Sheet1.cAccount - Globals.Sheet1.margins * 0.3) < 5000000) Then ' Arbitrary value, I can change.Lowest capital to make an operation
            Return True
        End If
        Return False
    End Function

    Public Function CanSellShort() As Boolean
        If (Globals.Sheet1.maxMargins - Globals.Sheet1.margins > 3000000) Then 'US$3Mil is arbitrary, I can change. Have in mind it changes my risk
            Return True
        End If
        Return False
    End Function

    Public Function MaxShortWithinConstraints(sym As String) As Double
        Dim q As Double = 0
        If (Globals.Sheet1.maxMargins - Globals.Sheet1.margins) < 1000000 Then ' arbitrary
            Return 0
        Else
            Dim maxIncreaseMargins = (Globals.Sheet1.maxMargins - Globals.Sheet1.margins)
            maxIncreaseMargins = maxIncreaseMargins - 3000000
            q = maxIncreaseMargins / FindPrice(sym, "Bid", CurrentDate)
            Return -Math.Truncate(q)
        End If
    End Function

    Public Function MaxBuy(sym As String) As Double
        Dim Q As Double = 0
        Dim availableCash As Double = 0
        If (Globals.Sheet1.cAccount - Globals.Sheet1.margins * 0.3) > 0 Then
            availableCash = (Globals.Sheet1.cAccount - Globals.Sheet1.margins) * 0.3
            availableCash = availableCash * 0.9
            Q = availableCash / FindPrice(sym, "Bid", CurrentDate)
            Return Math.Truncate(Q)
        Else
            Return 0
        End If
    End Function
End Module
