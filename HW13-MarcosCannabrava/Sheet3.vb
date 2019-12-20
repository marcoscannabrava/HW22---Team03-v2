
Public Class Sheet3

    Private Sub Sheet3_Startup() Handles Me.Startup
        Me.Activate()
        Application.ActiveWindow.DisplayHeadings = False
        Application.ActiveCell.Font.Size = 9
        InitialPositionLst.AutoSetDataBoundColumnHeaders = True
        TeamPortfolioLst.AutoSetDataBoundColumnHeaders = True
        ConfirmationTicketsLst.AutoSetDataBoundColumnHeaders = True
    End Sub

    Private Sub Sheet3_Shutdown() Handles Me.Shutdown

    End Sub


    Public Sub ClearListObjects()
        InitialPositionLst.DataBodyRange.Clear()
        TeamPortfolioLst.DataBodyRange.Clear()
        ConfirmationTicketsLst.DataBodyRange.Clear()

        InitialPositionLst.DataBodyRange.Font.Size = 9
        TeamPortfolioLst.DataBodyRange.Font.Size = 9
        ConfirmationTicketsLst.DataBodyRange.Font.Size = 9
    End Sub

    Public Sub ResetPortfolio()
        ClearPortfolio()
        AddToDBPortfolio("CAccount", GetInitialCAccount())
        Globals.Sheet1.ResetFinancialMetrics()
        Globals.Sheet1.ReCalcFinancialMetrics()
        Globals.Sheet1.DisplayFinancialMetrics()
        Globals.Sheet1.ResetTransactionData()
        Globals.Sheet1.DisplayTransactionData()
        Globals.Sheet5.InitializeCharts()


    End Sub

    Public Sub UploadPortfolioToDB()
        Dim newSymbol As String = ""
        Dim newUnits As String = ""

        If TeamPortfolioLst.IsSelected Then
            MessageBox.Show("Please click outside ListObject to confirm data entry")
            Exit Sub
        End If
        ClearPortfolio()

        For i As Integer = 1 To TeamPortfolioLst.DataBodyRange.Rows.Count()
            newSymbol = TeamPortfolioLst.DataBodyRange.Cells(i, 1).value
            newUnits = TeamPortfolioLst.DataBodyRange.Cells(i, 2).value
            If (newSymbol = Nothing) Or (newUnits = Nothing) Or (newSymbol = "") Or (newUnits = "") Then
                ' pass
            Else
                AddToDBPortfolio(newSymbol, newUnits)
            End If
        Next

        MessageBox.Show("Success! Portfolio updated on DB.", "Info")
        Globals.Sheet1.ResetFinancialMetrics()
        Globals.Sheet1.ReCalcFinancialMetrics()
        Globals.Sheet1.DisplayFinancialMetrics()
        Globals.Sheet1.ResetTransactionData()
        Globals.Sheet1.DisplayTransactionData()
    End Sub

    Public Sub UpdatePortfolio(transactionType As String, sym As String, qty As Double, TotTransactionValue As Double)
        'qty = Math.Abs(qty)
        Select Case transactionType
            Case "Buy"
                AddToDBPortfolio(sym, +qty)
            Case "Sell"
                AddToDBPortfolio(sym, -qty)
            Case "SellShort"
                AddToDBPortfolio(sym, -qty)
            Case "CashDiv"
                'pass
            Case "X-Put"
                'positionBT = GetPositionFromDB(sym)
                AddToDBPortfolio(GetUnderlier(sym), -qty)
                AddToDBPortfolio(sym, -qty)
            Case "X-Call"
                AddToDBPortfolio(GetUnderlier(sym), qty)
                AddToDBPortfolio(sym, -qty)
        End Select
        AddToDBPortfolio("CAccount", TotTransactionValue)
    End Sub

    Public Function CalcTPVatStart() As Double
        Return CalcIP(HTstartDate) + GetInitialCAccount()
    End Function

    Public Function CalcIP(targetDate As Date) As Double
        If excludeIPforTesting = True Then
            Return 0
        End If
        Dim tempSum As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double
        For Each myRow In myDataSet.Tables("InitialPositionTbl").Rows
            tempSymbol = myRow("Symbol").ToString().Trim
            tempUnits = myRow("Units")
            tempSum = tempSum + (tempUnits * CalcMTM(tempSymbol, targetDate))
        Next
        Return tempSum
    End Function

    Public Function CalcMTM(sym As String, targetDate As Date) As Double
        Dim mtm As Double = 0
        mtm = (FindPrice(sym, "ask", targetDate) + FindPrice(sym, "bid", targetDate)) / 2
        Return mtm
    End Function

    Public Function calcTaTPV(targetDate As Date) As Double
        Dim ts As TimeSpan = targetDate.Date - HTstartDate.Date
        Dim t As Double = ts.Days / 365.25
        Return Globals.Sheet1.TPVatStart * Math.Exp(Globals.Sheet1.iRate * t)
    End Function

    Public Function CalcTPV(targetDate As Date) As Double
        Return CalcIP(targetDate) + calcAP(targetDate) + Globals.Sheet1.cAccount + CalcInterestSLT(targetDate)
    End Function

    Public Function CalcAP(targetDate As Date) As Double
        Dim q As String = String.Format("Select * from {0}", TeamPortfolioTableName)
        RunQueryAndSaveResultsInDS(q, "TeamPFTbl")
        Dim tempAP As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double
        For Each myRow In myDataSet.Tables("TeamPFTbl").Rows
            tempSymbol = myRow("Symbol").ToString().Trim
            tempUnits = myRow("Units")
            If tempSymbol = "CAccount" Then
                'skip
            Else
                tempAP = tempAP + (tempUnits * CalcMTM(tempSymbol, targetDate))
            End If
        Next
        Return tempAP
    End Function

    Public Function CalcInterestSLT(toThisDay As Date) As Double
        Dim ts As TimeSpan = toThisDay.Date - LastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        Return Globals.Sheet1.cAccount * (Math.Exp(Globals.Sheet1.iRate * t) - 1)
    End Function

    Public Function CalcWeightedTE(targetDate As Date) As Double
        If Globals.Sheet1.TPV >= Globals.Sheet1.taTPV Then
            Return ((Globals.Sheet1.TPV - Globals.Sheet1.taTPV) / 4)
        Else
            Return (Globals.Sheet1.TPV - Globals.Sheet1.taTPV)
        End If
    End Function

    Public Function CalcMargins(targetDate As Date) As Double
        Dim tempMargin As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Integer
        If excludeIPforTesting = False Then
            For Each myRow In myDataSet.Tables("InitialPositionTbl").Rows
                tempSymbol = myRow("Symbol").ToString().Trim
                tempUnits = myRow("Units")
                If tempUnits < 0 Then
                    tempMargin = tempMargin + Math.Abs(tempUnits * CalcMTM(tempSymbol, targetDate))
                End If
            Next
        End If

        Dim q As String = String.Format("Select * from {0}", TeamPortfolioTableName)
        RunQueryAndSaveResultsInDS(q, "TeamPFTbl")
        For Each myRow In myDataSet.Tables("TeamPFTbl").Rows
            tempSymbol = myRow("Symbol").ToString().Trim
            tempUnits = myRow("Units")
            If (tempUnits) < 0 And (tempSymbol <> "CAccount") Then
                tempMargin = tempMargin + Math.Abs(tempUnits * CalcMTM(tempSymbol, targetDate))
            End If
        Next
        Return tempMargin
    End Function

    Public Function CalcEffectOfTransactionOnMargin(transactionType As String, symbol As String, qty As Integer) As Double
        Dim effect As Double = 0
        Dim positionBT As Integer = 0
        Dim underlierPositionBT As Integer = 0
        Select Case transactionType

            Case "Buy"
                positionBT = GetPositionFromDB(symbol)
                If positionBT < 0 Then
                    If qty >= Math.Abs(positionBT) Then
                        effect = positionBT * CalcMTM(symbol, CurrentDate)
                    End If
                    If qty < Math.Abs(positionBT) Then
                        effect = -qty * CalcMTM(symbol, CurrentDate)
                    End If
                End If

            Case "SellShort"
                effect = qty * CalcMTM(symbol, CurrentDate)

            Case "X-Call"
                positionBT = GetPositionFromDB(symbol)
                If positionBT < 0 Then
                    effect = positionBT * CalcMTM(symbol, CurrentDate)
                End If
                underlierPositionBT = GetPositionFromDB(GetUnderlier(symbol))
                If positionBT > 0 Then
                    If underlierPositionBT < 0 Then
                        If qty > Math.Abs(underlierPositionBT) Then
                            effect = effect + underlierPositionBT * CalcMTM(GetUnderlier(symbol), CurrentDate)
                        End If
                        If qty <= Math.Abs(underlierPositionBT) Then
                            effect = effect - qty * CalcMTM(GetUnderlier(symbol), CurrentDate)
                        End If
                    End If
                End If


            Case "X-Put"
                positionBT = GetPositionFromDB(symbol)
                If positionBT < 0 Then
                    effect = positionBT * CalcMTM(symbol, CurrentDate)
                End If
                underlierPositionBT = GetPositionFromDB(GetUnderlier(symbol))
                If positionBT < 0 Then
                    If underlierPositionBT < 0 Then
                        If qty > Math.Abs(underlierPositionBT) Then
                            effect = effect + underlierPositionBT * CalcMTM(GetUnderlier(symbol), CurrentDate)
                        End If
                        If qty <= Math.Abs(underlierPositionBT) Then
                            effect = effect - qty * CalcMTM(GetUnderlier(symbol), CurrentDate)
                        End If
                    End If
                End If
        End Select
        Return effect
    End Function
End Class
