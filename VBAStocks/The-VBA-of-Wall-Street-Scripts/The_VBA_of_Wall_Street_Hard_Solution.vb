Sub Stock_Market_Summary()

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

        ' Inserting vertical headers in column O, row 2-4
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Inserting horizontal headers in column P & Q
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Setting column width
        ws.Range("O:O").ColumnWidth = 21.5
        ws.Range("P:P").ColumnWidth = 14.25

        ' Determine the Last Row
        Last_Percentage_Change_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        Last_Total_Stock_Volume_Row = ws.Cells(Rows.Count, 12).End(xlUp).Row

        ' Extracting the Greatest % Increase
        Max_Percentage_Change = WorksheetFunction.Max(ws.Range("K2:K" & Last_Percentage_Change_Row))
        Max_Ticker = WorksheetFunction.Match(Max_Percentage_Change, ws.Range("K2:K" & Last_Percentage_Change_Row), 0)
        ws.Cells(2, 16).Value = ws.Cells(Max_Ticker + 1, 9)
        ws.Cells(2, 17).Value = FormatPercent(Max_Percentage_Change, 2)

        ' Extracting the Greatest % Decrease
        Min_Percentage_Change = WorksheetFunction.Min(ws.Range("K2:K" & Last_Percentage_Change_Row))
        Min_Ticker = WorksheetFunction.Match(Min_Percentage_Change, ws.Range("K2:K" & Last_Percentage_Change_Row), 0)
        ws.Cells(3, 16).Value = ws.Cells(Min_Ticker + 1, 9)
        ws.Cells(3, 17).Value = FormatPercent(Min_Percentage_Change, 2)

        ' Extracting the Greatest Total Volume
        Max_Total_Stock_Volume = WorksheetFunction.Max(ws.Range("L2:L" & Last_Total_Stock_Volume_Row))
        Max_Ticker = WorksheetFunction.Match(Max_Total_Stock_Volume, ws.Range("L2:L" & Last_Total_Stock_Volume_Row), 0)
        ws.Cells(4, 16).Value = ws.Cells(Max_Ticker + 1, 9)
        ws.Cells(4, 17).Value = Max_Total_Stock_Volume

        ' Setting column width
        ws.Range("Q:Q").EntireColumn.AutoFit

    Next ws

End Sub