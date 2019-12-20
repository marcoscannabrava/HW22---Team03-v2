Module DBProcedures
    Dim myConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection
    Dim myConnectionString As String = ""
    Dim myCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter
    Public myDataSet As DataSet = New DataSet


    Public Sub SetUpADOcomponents()
        Try
            myCommand.Connection = myConnection
            myDataAdapter.SelectCommand = myCommand
        Catch ex As Exception
            MessageBox.Show("Ops!SetUpADOcomponents failed: " + ex.Message)
        End Try

    End Sub

    Public Sub ConnectToDB(connString As String)
        Try
            myConnection.ConnectionString = connString
            myConnection.Open()
        Catch ex As Exception
            MessageBox.Show("Ops!ConnectToDB failed: " + ex.Message)
        End Try

    End Sub

    Public Sub DisconnectFromDB()
        Try
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show("Ops!DisconnectFromDB failed: " + ex.Message)
        End Try

    End Sub

    Public Sub RunQueryAndSaveResultsInDS(query As String, resultName As String)
        Try
            ClearTableInDS(resultName)
            myCommand.CommandText = query
            myDataAdapter.Fill(myDataSet, resultName)
        Catch ex As Exception
            MessageBox.Show("Ops!RunQueryAndSaveResultsInDS failed: " + ex.Message)
        End Try

    End Sub

    Public Sub ClearTableInDS(tableName As String)
        Try
            If myDataSet.Tables.Contains(tableName) Then
                myDataSet.Tables(tableName).Clear()
            End If
        Catch ex As Exception
            MessageBox.Show("Ops!ClearTableInDS failed: " + ex.Message)
        End Try

    End Sub

    Public Sub ExecuteNonQuery(query As String)
        Try
            myCommand.CommandText = query
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Ops!ExecuteNonQuery failed: " + ex.Message)
        End Try
    End Sub

    Public Sub ClearDS()
        myDataSet = New DataSet()
    End Sub

    Public Function GetADate(query As String) As Date
        Try
            myCommand.CommandText = query
            Return DateTime.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("Oops! GetADate failed! " + ex.Message + " 1/1/1 was returned.")
            Return DateTime.Parse("1/1/1")
        End Try
    End Function

    Public Function GetCurrentDate()
        Dim q As String = "select value from EnvironmentVariable where name = 'CurrentDate'"
        Try
            myCommand.CommandText = q
            Return Date.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("Could  not get the current date: " + ex.Message, "Error Message")
            Return Date.Parse("1/1/1")
        End Try
    End Function

    Public Function GetHistoricalPrice(symbol As String, askBid As String, targetDate As Date) As Double
        Dim q As String
        If IsAStock(symbol) Then
            q = String.Format("select {0} from StockMarket where ticker = '{1}' and date = '{2}'", _
                              askBid, symbol, targetDate.ToShortDateString())
        Else
            q = String.Format("select {0} from OptionMarket where ticker = '{1}' and date = '{2}'", _
                              askBid, symbol, targetDate.ToShortDateString())
        End If

        Try
            myCommand.CommandText = q
            Return Double.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("Could not get the historical price, " + ex.Message)
            Return 0
        End Try
    End Function

    Public Function GetInitialCAccount()
        Dim q As String = "select Value from EnvironmentVariable where Name = 'CAccount';"
        Try
            myCommand.CommandText = q
            Return Double.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("Cannot get the Capital Account. " & ex.Message, "Error message")
            Return 0
        End Try
    End Function

    Public Sub ClearPortfolio()
        Dim q As String = "delete from " & TeamPortfolioTableName
        ExecuteNonQuery(q)
    End Sub

    Public Sub AddToDBPortfolio(sym As String, qtyToAdd As Double)
        If qtyToAdd = 0 Then Exit Sub
        Dim newQty As Double
        Dim q As String
        newQty = GetPositionFromDB(sym) + qtyToAdd
        If newQty = 0 Then
            q = String.Format("delete from {0} where symbol = '{1}'", TeamPortfolioTableName, sym)
            ExecuteNonQuery(q)
        Else
            q = String.Format("delete from {0} where symbol = '{1}'", TeamPortfolioTableName, sym)
            ExecuteNonQuery(q)
            q = String.Format("insert into {0} values ('{1}', '{2}')", TeamPortfolioTableName, sym, newQty)
            ExecuteNonQuery(q)
        End If

    End Sub

    Public Function GetPositionFromDB(sym As String)
        Dim result As String
        Dim q As String = String.Format("select Units from {0} where Symbol = '{1}';", TeamPortfolioTableName, sym)
        Try
            myCommand.CommandText = q
            result = myCommand.ExecuteScalar()
            If result = "" Then
                Return 0
            Else
                Return Double.Parse(result)
            End If
        Catch ex As Exception
            MessageBox.Show("Could not get the position: " & ex.Message, "Error message")
            Return 0
        End Try
    End Function

    Public Function GetMaxMargins() As Double
        Dim q As String = "Select Value from EnvironmentVariable where Name = 'MaxMargins';"
        Try
            myCommand.CommandText = q
            Return Double.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("BZZZZ! I could not get you the maxMargins:" + ex.Message, "Error message")
            Return 0
        End Try
    End Function

    Public Function GetHTstartDate() As Date
        Dim q As String = "Select Value from EnvironmentVariable where Name = 'StartDate';"
        Try
            myCommand.CommandText = q
            Return Date.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("WOOOW! I could not get you the startdate" + ex.Message, "Error message")
            Return "1/1/1"
        End Try
    End Function

    Public Function GetIRate() As Double
        Dim q As String = "Select Value from EnvironmentVariable where name = 'RiskFreeRate';"
        Try
            myCommand.CommandText = q
            Return Double.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("I could not get you the iRate" + ex.Message, "Error message")
            Return 0
        End Try
    End Function

    Public Function GetLastTransactionDate() As Date
        Dim ltdate As String
        Dim userInput As String = "xyz"
        Dim q As String = String.Format("Select top(1) date from TransactionQueue where TeamId = '{0}' order by rowid;", TeamID)
        myCommand.CommandText = q
        ltdate = myCommand.ExecuteScalar()
        If ltdate = Nothing Or ltdate = "" Then
            ltdate = GetHTstartDate().ToShortDateString()
        End If
        Do
            userInput = InputBox("This date will be used as Last Transaction Date, unless you enter another date.", "Last transaction date", ltdate)
        Loop While Not IsDate(userInput)
        CurrentDate = Date.Parse(userInput)
        Return CurrentDate
    End Function

    Public Function GetMaxDate() As Date
        Dim q As String = "Select max(date) from StockMarket;"
        Try
            myCommand.CommandText = q
            Return Date.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("I could not get you the maxdate" + ex.Message, "Error Message")
            Return "1/1/1"
        End Try
    End Function

    Public Function NoData() As Boolean
        Dim q As String = "select top(1) ticker from stockmarket"
        myCommand.CommandText = q
        Dim result = myCommand.ExecuteScalar()
        If result = Nothing Or result = "" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetExpirationDate(symbol As String) as Datetime
        Dim q As String = String.Format("Select expiration from OptionMarket where symbol = '{0}'", symbol)
        Try
            myCommand.CommandText = q
            Return Date.Parse(myCommand.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show("Could  not get the current date: " + ex.Message, "Error Message")
            Return Date.Parse("1/1/1")
        End Try
    End Function

    Public Function IsInIP(Symbol As String) As Boolean

        If myDataSet.Tables.Contains("InitialPositionTbl") Then
            Dim myFilter As String = "Symbol = '" + Symbol + "'"
            Dim n As Integer = 0
            n = myDataSet.Tables("InitialPositionTbl").Select(myFilter).Count
            If n = 0 Then
                Return False
            Else
                Return True
            End If
        End If
        Return False
    End Function

    Public Function GetOptionType(symbol As String) As String
        Dim myFilter, temp As String
        Try
            myFilter = String.Format("Symbol = '{0}'", symbol)
            temp = myDataSet.Tables("OptionMarketForOneDayTbl").Select(myFilter).First.Item("Type").ToString()
            Return temp.Trim()
        Catch ex As Exception
            MessageBox.Show("Holy batskirt! I cannot find the type for " & symbol & ". " &
            ex.Message)
            Return ""
        End Try
    End Function
End Module



