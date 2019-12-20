Module DataSetProcedures

    Public Function IsAStock(symbol As String) As Boolean
        Dim answer As Boolean
        Dim myFilter As String = "Ticker = '" & symbol & "'"
        Dim n As Integer = 0
        n = myDataSet.Tables("TickerTbl").Select(myFilter).Count
        If n = 0 Then
            answer = False
        Else
            answer = True
        End If
        Return answer
    End Function

    Public Function GetTransactionCostCoefficient(SecurityType As String, TransactionType As String) As Double
        Dim myFilter As String
        Dim temp As String
        myFilter = "SecurityType = '" & SecurityType & "' and TransactionType = '" & TransactionType & "'"
        temp = myDataSet.Tables("TransactionCostTbl").Select(myFilter).First.Item("CostCoeff").ToString()
        Return Double.Parse(temp)
    End Function

    Public Function FindPrice(symbol As String, askBid As String, targetDate As Date) As Double
        Try
            If targetDate.Date <> CurrentDate Then
                Return GetHistoricalPrice(symbol, askBid, targetDate)
            End If
            Dim price As Double = 0
            Dim myFilter As String
            Dim temp As String
            If targetDate.DayOfWeek = DayOfWeek.Saturday Then
                targetDate = targetDate.AddDays(-1)
            End If
            If targetDate.DayOfWeek = DayOfWeek.Sunday Then
                targetDate = targetDate.AddDays(-2)
            End If
            If IsAStock(symbol) Then
                myFilter = String.Format("Ticker = '{0}' and date = '{1}'", symbol, targetDate.ToShortDateString())
                temp = myDataSet.Tables("StockMarketForOneDayTbl").Select(myFilter).First.Item(askBid).ToString()
            Else
                myFilter = String.Format("Symbol = '{0}' and date = '{1}'", symbol, targetDate.ToShortDateString())
                temp = myDataSet.Tables("OptionMarketForOneDayTbl").Select(myFilter).First.Item(askBid).ToString()
            End If

            Return Double.Parse(temp)
        Catch ex As Exception
            MessageBox.Show("Cannot find the price for " & symbol & "." & ex.Message)
            Return 0
        End Try

    End Function

    Public Function FindDividend(symbol As String, targetDate As Date) As Double
        Try
            If targetDate.Date <> CurrentDate Then
                'Return GetHistoricalDividend(symbol, targetDate)
            End If
            Dim dividend As Double = 0
            Dim myFilter As String
            Dim temp As String
            If targetDate.DayOfWeek = DayOfWeek.Saturday Then
                targetDate = targetDate.AddDays(-1)
            End If
            If targetDate.DayOfWeek = DayOfWeek.Sunday Then
                targetDate = targetDate.AddDays(-2)
            End If
            If IsAStock(symbol) Then
                myFilter = String.Format("Ticker = '{0}' and date = '{1}'", symbol, targetDate.ToShortDateString())
                temp = myDataSet.Tables("StockMarketForOneDayTbl").Select(myFilter).First.Item("Dividend").ToString()
                Return Double.Parse(temp)
            Else
                Return 0
            End If
            'Return Double.Parse(temp)
        Catch ex As Exception
            MessageBox.Show("Cannot find the dividend for " & symbol & "." & ex.Message)
            Return 0
        End Try

    End Function

    Public Function GetStrike(symbol As String) As Double
        Dim myFilter, temp As String
        Try
            myFilter = String.Format("Symbol = '{0}'", symbol)
            temp = myDataSet.Tables("OptionMarketForOneDayTbl").Select(myFilter).First.Item("Strike").ToString()
            Return Double.Parse(temp)
        Catch ex As Exception
            MessageBox.Show("I cannot find the strike for" + symbol + "." + ex.Message)
            Return 0
        End Try
    End Function

    Public Function GetUnderlier(symbol As String) As String
        Dim myFilter, temp As String
        Try
            myFilter = String.Format("Symbol = '{0}'", symbol)
            temp = myDataSet.Tables("OptionMarketForOneDayTbl").Select(myFilter).First.Item("Underlier").ToString()
            Return temp.Trim()
        Catch ex As Exception
            MessageBox.Show("I cannot find the underlier for" + symbol + "." + ex.Message)
            Return ""
        End Try
    End Function
End Module
