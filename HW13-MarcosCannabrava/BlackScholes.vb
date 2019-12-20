Module BlackScholes
    Public Function CalcFamilyDelta(tkr As String) As Double
        Dim tempSum As Double = 0
        Dim delta As Double = 0
        Dim sym As String
        tkr = tkr.Trim()

        'Acquired Postion
        Dim q = "select * from " & TeamPortfolioTableName
        RunQueryAndSaveResultsInDS(q, "TeamPortfolioTbl")
        For Each dr As DataRow In myDataSet.Tables("TeamPortfolioTbl").Rows
            sym = dr("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, CurrentDate)
                tempSum = tempSum + delta * dr("units")
            End If
        Next

        'Initial Postion
        If excludeIPforTesting = True Then
            '
        Else
            For Each dr As DataRow In myDataSet.Tables("InitialPositionTbl").Rows
                sym = dr("Symbol").ToString().Trim()
                If IsInTheFamily(sym, tkr) Then
                    delta = CalcDelta(sym, CurrentDate)
                    tempSum = tempSum + delta * dr("Units")
                End If
            Next
        End If
        Return tempSum
    End Function

    Public Function IsInTheFamily(sym As String, familyTicker As String) As Boolean
        If sym = "CAccount" Then
            Return False
        End If
        If IsAStock(sym) Then 'stock
            If sym = familyTicker Then
                Return True
            Else
                Return False
            End If
        Else 'option
            If GetUnderlier(sym) = familyTicker Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Function CalcDelta(symbol As String, targetDate As Date) As Double
        Dim sigma As Double
        Dim K As Double
        Dim S As Double
        Dim r As Double = Globals.Sheet1.iRate
        Dim t As Double
        Dim ts As TimeSpan
        Dim underlier As String
        Dim d1 As Double


        If symbol = "CAccount" Then
            Return 0
        End If
        If IsAStock(symbol) Then
            Return 1
        End If
        If targetDate.Date >= GetExpirationDate(symbol).date Then
            Return 0
        End If
        If FindPrice(symbol, "Ask", targetDate) = 0 Then
            Return 0
        End If


        underlier = GetUnderlier(symbol)
        sigma = GetVolatility(underlier)
        K = GetStrike(symbol)
        S = Globals.Sheet3.CalcMTM(underlier, targetDate)
        ts = GetExpirationDate(symbol).Date - targetDate.Date
        t = ts.Days / 365.25
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))
        If GetOptionType(symbol) = "Call" Then
            Return Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True)
        End If
        If GetOptionType(symbol) = "Put" Then
            Return (Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True) - 1)
        End If
        Return 0
    End Function

    Public Function GetVolatility(tkr As String) As Double
        For i = 1 To 12
            If Globals.Sheet1.UnderlierRange.Cells(i, 1).Value.Trim() = tkr.Trim() Then
                Return Volatilities(i - 1)
            End If
        Next
        Return 0
    End Function

End Module
