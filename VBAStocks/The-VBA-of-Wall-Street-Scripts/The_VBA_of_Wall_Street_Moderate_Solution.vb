Sub Yearly_and_Percent_Change()

    ' Set an initial variable for holding the Ticker symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for holding the Opening Price
    Dim Opening_Price As Double

    ' Set an initial variable for holding the Closing Price
    Dim Closing_Price As Double

    ' Set an initial variable for holding the Yearly Change in Price
    Dim Yearly_Change_In_Price As Double

    ' Set an initial variable for holding the Percent Change in Price
    Dim Percent_Change_In_Price As Double

    ' Set an initial variable for holding the Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

        ' Loop through all worksheets
        For Each ws In Worksheets

        ' Inserting headers for column I, J, K & L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Setting column width
        ws.Range("I:K").ColumnWidth = 14.25
        ws.Range("L:L").EntireColumn.AutoFit

        ' Keep track of the location for each Ticker symbol
        Dim Ticker_Row As Integer
        Ticker_Row = 2

        ' Keep track of the location for each Opening Price
        Dim Opening_Price_Row As Long
        Opening_Price_Row = 2

        ' Determine the Last Row
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all Ticker symbols
        For i = 2 To Last_Row

            ' Determine if the Ticker symbol is the same or different from previous
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value

                ' Set the Opening Price
                Opening_Price = ws.Cells(Opening_Price_Row, 3).Value

                ' Set the Closing Price
                Closing_Price = ws.Cells(i, 6).Value

                ' Calculate Yearly Change in Price
                Yearly_Change_In_Price = Closing_Price - Opening_Price

                ' Calculate Percent Change in Price
                If Opening_Price = 0 Then
                    Percent_Change_In_Price = 0

                Else 
                    Percent_Change_In_Price = Yearly_Change_In_Price / Opening_Price

                End If

                 ' Total the Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                ' Print the Ticker symbol in Column I
                ws.Range("I" & Ticker_Row).Value = Ticker_Symbol

                ' Print the Yearly Change in Price in Column J
                ws.Range("J" & Ticker_Row).Value = Yearly_Change_In_Price

                    'Conditional formatting highlighting positive change in green and negative change in red
                    If Yearly_Change_In_Price > 0 Then
                    ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4

                    Else
                    ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3

                    End If

                ' Print the Percent Change in Price in Column K
                ws.Range("K" & Ticker_Row).Value = FormatPercent(Percent_Change_In_Price, 2)

                ' Print the Total Stock Volume in Column L
                ws.Range("L" & Ticker_Row).Value = Total_Stock_Volume

                ' Add new row to the Ticker column
                Ticker_Row = Ticker_Row + 1

                Opening_Price_Row = i + 1

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