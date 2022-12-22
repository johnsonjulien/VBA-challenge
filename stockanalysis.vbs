Sub testing_loop()
'For each WS in Worksheets
'ws.

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"

    Range("L1").Value = "Total Stock Volume"

    Dim Ticker As String
    Dim Open_Price, Close_Price As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As LongLong

    ' Set an initial variable for holding the brand name (Ticker)
    ' Set an initial variable for holding the total per credit card brand (Yearly_Change)
    ' Set an initial variable for holding the total per credit card brand (Percent_Change)
    ' Set an initial variable for holding the total per credit card brand (Stock_Volume
    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Total_Stock_Volume = 0

    Open_Price = Cells(2, "C").Value


    ' Loop through all credit card purchases(Tickers)
    For i = 2 To 22771

        Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value

        ' Check if we are still within the same credit card brand, if it is not...(Ticker)
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Brand name (Ticker)
            Ticker = Cells(i, 1).Value

            ' Add to the Brand Total(Yearly_Change)
            Close_Price = Cells(i, "F").Value
            Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price


            If Range("J" & Summary_Table_Row).Value > 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 'green'

            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 'red'

            End If
            'Add to the Brand Total(Percent_Change)
            If Open_Price <> 0 Then

                Range("K" & Summary_Table_Row).Value = FormatPercent(Range("J" & Summary_Table_Row).Value / Open_Price, 2)
            Else

                Range("K" & Summary_Table_Row).Value = Null

            End If


            'Print the Credit Card Brand in the Summary Table (Print Ticker)
            Range("I" & Summary_Table_Row).Value = Ticker

            ' Print the Brand Amount to the Summary Table (Print Stock Volume)
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1


            ' If the cell immediately following a row is the same brand...


            ' Add to the Stock_Volume
            Total_Stock_Volume = 0
            Open_Price = Cells(i + 1, "C").Value

        End If

    Next i

    'The ticker symbol. Loop through the columns
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    'Make sure to use conditional formatting that will highlight positive change in green and negative change in red.





'Next ws
End Sub