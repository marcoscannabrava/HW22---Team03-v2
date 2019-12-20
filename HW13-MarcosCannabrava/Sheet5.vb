
Public Class Sheet5

    Private Sub Sheet5_Startup() Handles Me.Startup
        Application.ActiveWindow.DisplayHeadings = False
        Application.ActiveCell.Font.Size = 9
        TEtoChartLST.AutoSetDataBoundColumnHeaders = True
        SecurityToChartLST.AutoSetDataBoundColumnHeaders = True
    End Sub

    Private Sub Sheet5_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles StockToChartCBox.SelectedIndexChanged
        Dim sql As String
        sql = String.Format("Select date, bid, ask from StockMarket where ticker = '{0}' order by date", StockToChartCBox.SelectedItem.trim())
        RunQueryAndSaveResultsInDS(sql, "DataToChartTbl")
        SecurityToChartLST.DataSource = myDataSet.Tables("DataToChartTbl")
        SecurityChart.SetSourceData(SecurityToChartLST.Range)

        SecurityChart.ChartTitle.Text = "Daily Closings for " + StockToChartCBox.SelectedItem
        Dim y As Excel.Axis = SecurityChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate((findMinBid("DataToChartTbl") * 0.9 / 10) * 10)
    End Sub

    Public Function findMinBid(tablename As String) As Double
        Dim temp As String = "0"
        temp = myDataSet.Tables(tablename).Compute("min(Bid)", "").ToString()
        If temp = "" Then
            temp = "0"
        End If
        Return Double.Parse(temp)
    End Function


    Public Sub ClearListObjects()
        Try
            SecurityToChartLST.DataBodyRange.Clear()
            TEtoChartLST.DataBodyRange.Clear()

            SecurityToChartLST.DataBodyRange.Font.Size = 9
            TEtoChartLST.DataBodyRange.Font.Size = 9
        Catch ex As Exception
            '
        End Try
        
    End Sub

    Public Sub InitializeCharts()
        ClearTEdataTbl()
        InitializeSecurityChart()
        InitializeTEchart()
    End Sub

    Public Sub ClearTEdataTbl() 'CREATING THE TABLE WITH THE DATA
        If myDataSet.Tables.Contains("TEdataTbl") Then
            myDataSet.Tables("TEdataTbl").Clear()
        Else
            Dim t As DataTable = myDataSet.Tables.Add("TEdataTbl")
            t.Columns.Add("Date", GetType(Date))
            t.Columns.Add("TPV", GetType(Double))
            t.Columns.Add("TaTPV", GetType(Double))
        End If
        TEtoChartLST.DataSource = myDataSet.Tables("TEdataTbl")
    End Sub

    Public Sub InitializeSecurityChart()

        Dim q As String = "select ticker from StockTicker order by ticker"
        RunQueryAndSaveResultsInDS(q, "TickerTbl")
        StockToChartCBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("TickerTbl").Rows
            StockToChartCBox.Items.Add(dr("Ticker").ToString.Trim())
        Next
        StockToChartCBox.Text = "Select"

        q = "select symbol from OptionSymbol order by symbol"
        RunQueryAndSaveResultsInDS(q, "SymbolTbl")
        OptionToChartCBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("SymbolTbl").Rows
            OptionToChartCBox.Items.Add(dr("Symbol").ToString.Trim())
        Next
        OptionToChartCBox.Text = "Select"

        SecurityChart.ChartType = Excel.XlChartType.xlLine
        SecurityChart.ChartStyle = 42
        SecurityChart.ApplyLayout(1)

        Dim y As Excel.Axis = SecurityChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$###.00"

        Dim x As Excel.Axis = SecurityChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"
    End Sub


    Public Sub InitializeTEchart() 'CREATING THE CHART
        Globals.Sheet1.Chart_1.ChartType = Excel.XlChartType.xlLine
        Globals.Sheet1.Chart_1.ChartStyle = 48
        Globals.Sheet1.Chart_1.ApplyLayout(12)
        Globals.Sheet1.Chart_1.HasTitle = False

        Dim y As Excel.Axis = Globals.Sheet1.Chart_1.Axes(Excel.XlAxisType.xlCategory)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"

        Dim x As Excel.Axis = Globals.Sheet1.Chart_1.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"
    End Sub

    Private Sub OptionToChartCBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OptionToChartCBox.SelectedIndexChanged
        Dim sql As String
        sql = String.Format("Select date, bid, ask from OptionMarket where symbol = '{0}' order by date", OptionToChartCBox.SelectedItem.trim())
        RunQueryAndSaveResultsInDS(sql, "DataToChartTbl")
        SecurityToChartLST.DataSource = myDataSet.Tables("DataToChartTbl")
        SecurityChart.SetSourceData(SecurityToChartLST.Range)

        SecurityChart.ChartTitle.Text = "Daily Closings for " + OptionToChartCBox.SelectedItem
        Dim y As Excel.Axis = SecurityChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate((findMinBid("DataToChartTbl") * 0.9 / 10) * 10)
    End Sub

    Public Sub ProjectPortfolio()
        While CurrentDate < GetMaxDate()
            CurrentDate = CurrentDate.AddDays(1)
            DailyRoutine()
        End While
    End Sub

    Public Sub AddRowToTETable(targetDate As Date) 'POPULATING THE TABLE
        Dim tempTPV, tempTaTPV
        Dim dr As DataRow
        Dim filter As String = String.Format("Date = '{0}'", targetDate.ToShortDateString())
        dr = myDataSet.Tables("TEdataTbl").Select(filter).FirstOrDefault
        If IsNothing(dr) Then
            tempTPV = Globals.Sheet3.CalcTPV(targetDate)
            tempTaTPV = Globals.Sheet3.calcTaTPV(targetDate)
            myDataSet.Tables("TEdataTbl").Rows.Add(targetDate, tempTPV, tempTaTPV)
        Else
            dr(1) = Globals.Sheet1.TPV
            dr(2) = Globals.Sheet1.taTPV
        End If
        Dim y As Excel.Axis = Globals.Sheet1.Chart_1.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate((findMinTPV("TEdataTbl") * 0.9) / 100000) * 100000


        Globals.Sheet1.Chart_1.SetSourceData(Globals.Sheet5.TEtoChartLST.Range)
    End Sub

    Public Function findMinTPV(tablename As String) As Double
        Dim temp As String = "0"
        temp = myDataSet.Tables(tablename).Compute("min(TPV)", "").ToString()
        If temp = "" Then
            temp = "0"
        End If
        Return Double.Parse(temp)
    End Function

   
End Class
