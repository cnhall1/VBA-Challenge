Sub Ticker_Symbol_and_Total_Stock_Volume()

    ' Set an initial variable for holding the Ticker symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for holding the Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

        ' Loop through all worksheets
        For Each ws In Worksheets

        ' Inserting headers for column I & J
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"

        ' Setting column width
        ws.Range("I:I").ColumnWidth = 14.25
        ws.Range("J:J").EntireColumn.AutoFit

        ' Keep track of the location for each Ticker symbol
        Dim Ticker_Row As Integer
        Ticker_Row = 2

        ' Determine the Last Row
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all Ticker symbols
        For i = 2 To Last_Row

            ' Determine if the Ticker symbol is the same or different from previous
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' Total the Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                ' Print the Ticker symbol in Column I
                ws.Range("I" & Ticker_Row).Value = Ticker_Symbol

                ' Print the Total Stock Volume in Column J
                ws.Range("J" & Ticker_Row).Value = Total_Stock_Volume

                ' Add new row to the Ticker column
                Ticker_Row = Ticker_Row + 1

                ' Reset the Stock Volume Total
                Total_Stock_Volume = 0
                
            ' If the cell immediately following a row is the same Ticker symbol
            Else

                ' Total the Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            End If

        Next i

    Next ws

End Sub