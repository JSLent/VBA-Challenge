'Module 2 Challenge Stock Analysis
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once

Sub StockDataAnalysis()
    'Assign variables As value
    Dim ws As Worksheet
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim StartPrice As Double
    Dim EndPrice As Double
    Dim LastRow As Long
    Dim SummaryRow As Integer

    'Loop through each worksheet
    For Each ws In Worksheets
        SummaryRow = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Define Ranges and input column header for output
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Range("O2").Value = Array("Greatest % Increase")
        ws.Range("O3").Value = Array("Greatest % Decrease")
        ws.Range("O4").Value = Array("Greatest Total Volume")
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Assign variables for additional functionality task and set starting values for each
        Dim greatestIncrease As Double: greatestIncrease = 0
        Dim greatestDecrease As Double: greatestDecrease = 0
        Dim greatestVolume As Double: greatestVolume = 0
        Dim tickerIncrease As String: tickerIncrease = ""
        Dim tickerDecrease As String: tickerDecrease = ""
        Dim tickerVolume As String: tickerVolume = ""

        ' Loop through all rows
        For i = 2 To LastRow
            'Check for Ticker, EndPrice, YearlyChange, PercentChange, TotalVolume values for summary table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Assign value ticker
                Ticker = ws.Cells(i, 1).Value
                'Assign value to endPrice
                EndPrice = ws.Cells(i, 6).Value
                'Calculate yearlyChange
                YearlyChange = EndPrice - StartPrice
                'Calculate percent of yearlyChange and round down
                PercentChange = Round((YearlyChange / StartPrice) * 100, 2)
                'Calculate TotalValue
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ' Create summary table with output values
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Conditional formatting for YearlyChange
                If YearlyChange > 0 Then
                    'Green
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    'Red
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                ' Check for greatest increase, decrease, and volume
                If PercentChange > greatestIncrease Then
                    greatestIncrease = PercentChange
                    tickerIncrease = Ticker
                ElseIf PercentChange < greatestDecrease Then
                    greatestDecrease = PercentChange
                    tickerDecrease = Ticker
                End If

                'Check for Ticker with highest Volume
                If TotalVolume > greatestVolume Then
                    greatestVolume = TotalVolume
                    tickerVolume = Ticker
                End If
                'Move to next row and reset
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
            Else
                'Increase TotalVolume by value in Volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If

            ' Set the start price for the next stock
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Set StartPrice to value in Open
                StartPrice = ws.Cells(i, 3).Value
            End If
        Next i

        ' Output the greatest values and their corresponding tickers
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(4, 17).Value = greatestVolume
    Next ws

End Sub
