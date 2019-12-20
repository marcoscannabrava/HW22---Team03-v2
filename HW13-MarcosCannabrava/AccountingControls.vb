Module AccountingControls
    Public Function IsTradeValid() As Boolean
        If Globals.Sheet1.tSymbol = "" Or Globals.Sheet1.tSymbol = "Stock" Or Globals.Sheet1.tSymbol = "Option" Then
            Return False
        End If
        If Globals.Sheet1.tType = "" Or Globals.Sheet1.tType = "Select" Then
            Return False
        End If

        If Globals.Sheet1.tQty <= 0 And (Globals.Sheet1.tType <> "CashDiv") And
            (Globals.Sheet1.tType <> "X-Put") And
            (Globals.Sheet1.tType <> "X-Call") Then
            Return False
        End If
        If Globals.Sheet1.tPrice = 0 Then
            Return False
        End If
        If Globals.Sheet1.tType = "Sell" And (Globals.Sheet1.tQty > GetPositionFromDB(Globals.Sheet1.tSymbol)) Then
            Return False
        End If
        If Globals.Sheet1.tType = "SellShort" And (GetPositionFromDB(Globals.Sheet1.tSymbol) > 0) Then
            Return False
        End If

        Return True
    End Function
End Module
