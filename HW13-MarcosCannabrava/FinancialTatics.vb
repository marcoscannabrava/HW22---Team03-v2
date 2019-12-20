Module FinancialTatics

    Public Function CheckSellStock(tkr As String) As Boolean
        If IsInIP(tkr) Then
            Return False
        End If
        Dim position As Double = GetPositionFromDB(tkr)
        Dim qtyToHedge As Double = 0
        If position > 0 Then
            qtyToHedge = CalcSignedQtyNeededToHedge(tkr, tkr)
            If qtyToHedge = 0 Then
                Return False
            End If
            If position >= Math.Abs(qtyToHedge) Then 'you have more than needed
                suggestedSignedQty = qtyToHedge
            Else
                suggestedSignedQty = -position 'sell all you have
            End If
            suggestedSymbol = tkr
            suggestedAction = "Sell"
            Return True
        Else 'position is negative, then return false
            Return False
        End If
    End Function

    Public Function CheckSellCall(underlier As String) As Boolean
        Dim sym As String = ""
        For Each dr As DataRow In myDataSet.Tables(TeamPortfolioTableName).rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                'skip
            Else
                If (GetOptionType(sym) = "Call") And (GetUnderlier(sym) = underlier) Then
                    Dim position As Double = dr("Units")
                    Dim qtyToHedge As Double = 0
                    If position > 0 Then
                        qtyToHedge = CalcSignedQtyNeededToHedge(underlier, sym)
                        If qtyToHedge <> 0 Then
                            If position >= Math.Abs(qtyToHedge) Then
                                suggestedSignedQty = qtyToHedge
                            Else
                                suggestedSignedQty = -position 'sell all you have
                            End If
                            suggestedSymbol = sym
                            suggestedAction = "Sell"
                            Return True
                        End If
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckBuyBackPut(underlier As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        Dim sym As String = ""
        For Each dr As DataRow In myDataSet.Tables(TeamPortfolioTableName).rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                'skip
            Else
                If (GetOptionType(sym) = "Put") And (GetUnderlier(sym) = underlier) Then
                    Dim position As Double = dr("Units")
                    If (position < 0) Then
                        Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                        If Math.Abs(position) < qtyToHedge Then
                            suggestedSignedQty = -position 'buy back all that you have
                        Else
                            suggestedSignedQty = qtyToHedge
                        End If
                        'how much can you afford?
                        Dim maxAffordableQty As Double = MaxShortBuyBack(sym)
                        If maxAffordableQty < suggestedSignedQty Then
                            suggestedSignedQty = maxAffordableQty
                        End If
                        suggestedSymbol = sym
                        suggestedAction = "Buy"
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckSellShortCall(underlier As String) As Boolean
        If CanSellShort() Then
            Dim sym As String = ""
            For Each partialSymbol As String In {"_COCTA", "_COCTB", "_COCTC", "_COCTD", "_COCTE"}
                sym = underlier + partialSymbol
                If Not IsInIP(sym) Then
                    Dim position As Double = GetPositionFromDB(sym)
                    If (position <= 0) Then
                        Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                        If qtyToHedge <> 0 Then
                            Dim maxShortQty As Double = MaxShortWithinConstraints(sym)
                            If Math.Abs(qtyToHedge) > Math.Abs(maxShortQty) Then
                                suggestedSignedQty = maxShortQty
                            Else
                                suggestedSignedQty = qtyToHedge
                            End If
                        'how much can you afford?
                            suggestedSymbol = sym
                            suggestedAction = "SellShort"
                        Return True
                        End If
                    End If
                End If
            Next
        End If
        Return False
    End Function

    Public Function CheckSellShortStock(ticker As String) As Boolean
        If CanSellShort() Then
            If Not IsInIP(ticker) Then
                Dim position As Double = GetPositionFromDB(ticker)
                If (position <= 0) Then
                    Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(ticker, ticker)
                    If qtyToHedge <> 0 Then
                        Dim maxShortQty As Double = MaxShortWithinConstraints(ticker)
                        If Math.Abs(qtyToHedge) > Math.Abs(maxShortQty) Then
                            suggestedSignedQty = maxShortQty
                        Else
                            suggestedSignedQty = qtyToHedge
                        End If
                        suggestedSymbol = ticker
                        suggestedAction = "SellShort"
                        Return True
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Public Function CheckBuyPut(underlier As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        Dim sym As String = ""
        For Each partialSymbol As String In {"_POCTE", "_POCTD", "_POCTC", "_POCTB", "_POCTA"}
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                Dim position As Double = GetPositionFromDB(sym)
                If (position >= 0) Then
                    Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                    If qtyToHedge <> 0 Then
                        Dim maxAffordableQty As Double = MaxBuy(sym)
                        If maxAffordableQty < qtyToHedge Then
                            suggestedSignedQty = maxAffordableQty
                        Else
                            suggestedSignedQty = qtyToHedge
                        End If
                    suggestedSymbol = sym
                    suggestedAction = "Buy"
                    Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckSellPut(underlier As String) As Boolean
        Dim sym As String = ""
        For Each dr As DataRow In myDataSet.Tables(TeamPortfolioTableName).rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                'skip
            Else
                If (GetOptionType(sym) = "Put") And (GetUnderlier(sym) = underlier) Then
                    Dim position As Double = GetPositionFromDB(sym)
                    If position > 0 Then
                        Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                        If qtyToHedge <> 0 Then
                            If position >= Math.Abs(qtyToHedge) Then
                                suggestedSignedQty = qtyToHedge
                            Else
                                suggestedSignedQty = -position 'sell all you have
                            End If
                            suggestedSymbol = sym
                            suggestedAction = "Sell"
                            Return True
                        End If
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckSellShortPut(underlier As String) As Boolean
        If CanSellShort() Then
            Dim sym As String = ""
            For Each partialSymbol As String In {"_POCTE", "_POCTD", "_POCTC", "_POCTB", "_POCTA"}
                sym = underlier + partialSymbol
                If Not IsInIP(sym) Then
                    Dim position As Double = GetPositionFromDB(sym)
                    If (position <= 0) Then
                        Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                        If qtyToHedge <> 0 Then
                            Dim maxShortQty As Double = MaxShortWithinConstraints(sym)
                            If Math.Abs(qtyToHedge) > Math.Abs(maxShortQty) Then
                                suggestedSignedQty = maxShortQty
                            Else
                                suggestedSignedQty = qtyToHedge
                            End If
                            suggestedSymbol = sym
                            suggestedAction = "SellShort"
                            Return True
                        End If
                    End If
                End If
            Next
        End If
        Return False
    End Function

    Public Function CheckBuyBackCall(underlier As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        Dim sym As String = ""
        For Each dr As DataRow In myDataSet.Tables(TeamPortfolioTableName).rows
            sym = dr("Symbol").ToString().Trim()
            If IsAStock(sym) Or sym = "CAccount" Then
                'skip
            Else
                If (GetOptionType(sym) = "Call") And (GetUnderlier(sym) = underlier) Then
                    Dim position As Double = dr("Units")
                    If (position < 0) Then
                        Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                        If qtyToHedge <> 0 Then
                            If Math.Abs(position) < qtyToHedge Then
                                suggestedSignedQty = -position
                            Else
                                suggestedSignedQty = qtyToHedge
                            End If
                            Dim maxAffordableQty As Double = MaxShortBuyBack(sym)
                            If maxAffordableQty < suggestedSignedQty Then
                                suggestedSignedQty = maxAffordableQty
                            End If
                            suggestedSymbol = sym
                            suggestedAction = "Buy"
                            Return True
                        End If
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckBackStock(ticker As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        If Not IsInIP(ticker) Then
            Dim position As Double = GetPositionFromDB(ticker)
            If (position < 0) Then
                Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(ticker, ticker)
                If qtyToHedge <> 0 Then
                    If Math.Abs(position) < qtyToHedge Then
                        suggestedSignedQty = -position
                    Else
                        suggestedSignedQty = qtyToHedge
                    End If
                    Dim maxAffordableQty As Double = MaxShortBuyBack(ticker)
                    If maxAffordableQty < suggestedSignedQty Then
                        suggestedSignedQty = maxAffordableQty
                    End If
                    suggestedSymbol = ticker
                    suggestedAction = "Buy"
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Public Function CheckBuyCall(ByVal underlier As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        Dim sym As String = ""
        For Each partialSymbol As String In {"_COCTA", "_COCTB", "_COCTC", "_COCTD", "_COCTE"}
            sym = underlier + partialSymbol
            If Not IsInIP(sym) Then
                Dim position As Double = GetPositionFromDB(sym)
                If (position >= 0) Then
                    Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(underlier, sym)
                    If qtyToHedge <> 0 Then
                        Dim maxAffordableQty As Double = MaxBuy(sym)
                        If maxAffordableQty < qtyToHedge Then
                            suggestedSignedQty = maxAffordableQty
                        Else
                            suggestedSignedQty = qtyToHedge
                        End If
                        suggestedSymbol = sym
                        suggestedAction = "Buy"
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Function CheckBuyStock(ByVal ticker As String) As Boolean
        If AvailableCashIsLow() Then
            Return False
        End If
        If Not IsInIP(ticker) Then
            Dim position As Double = GetPositionFromDB(ticker)
            If (position >= 0) Then
                Dim qtyToHedge As Double = CalcSignedQtyNeededToHedge(ticker, ticker)
                If qtyToHedge <> 0 Then
                    Dim maxAffordableQty As Double = MaxBuy(ticker)
                    If maxAffordableQty < qtyToHedge Then
                        suggestedSignedQty = maxAffordableQty
                    Else
                        suggestedSignedQty = qtyToHedge
                    End If
                    suggestedSymbol = ticker
                    suggestedAction = "Buy"
                    Return True
                End If
            End If
        End If
        Return False
    End Function


End Module
