Sub TickerOnAllWorksheets()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Call the Ticker subroutine for each worksheet
        Ticker ws
    Next ws
End Sub

Sub Ticker(ws As Worksheet)
    ' Set column values
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"

    ' Set initial variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim First_Open_Price As Double ' Variable to store the first day's opening price
    Dim Close_Price As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim lastRow As Long
    Dim maxPercentIncrease As Double
    Dim minPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim minPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String

    ' Initialize summary table row
    Summary_Table_Row = 2

    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row of data
    For i = 2 To lastRow
        ' Get the ticker symbol
        Ticker = ws.Cells(i, 1).Value

        ' Check if it's a new ticker symbol
        If ws.Cells(i + 1, 1).Value <> Ticker Then
            ' Get the closing price and volume
            Close_Price = ws.Cells(i, 6).Value

            ' Calculate yearly change (closing price - first day's opening price), percent change, and stock volume
            Yearly_Change = Close_Price - First_Open_Price
            Percent_Change = Yearly_Change / First_Open_Price
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            ' Print ticker, yearly change, percent change, and total volume in the summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "0.00%")
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

            'Add conditional formatting
            If Yearly_Change < 0 Then
            ' Red for neg
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            ' Green for pos
            ElseIf Yearly_Change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If

            If Percent_Change < 0 Then
            ' Red for neg
            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            ' Green for pos
            ElseIf Percent_Change > 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If

            ' Update maximum percent increase
            If Percent_Change > maxPercentIncrease Then
                maxPercentIncrease = Percent_Change
                maxPercentIncreaseTicker = Ticker
            End If

            ' Update minimum percent decrease
            If Percent_Change < minPercentDecrease Then
                minPercentDecrease = Percent_Change
                minPercentDecreaseTicker = Ticker
            End If

            ' Update maximum total volume
            If Total_Stock_Volume > maxTotalVolume Then
                maxTotalVolume = Total_Stock_Volume
                maxTotalVolumeTicker = Ticker
            End If

            ' Move to the next row in the summary table
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset variables for the next ticker
            Yearly_Change = 0
            First_Open_Price = 0
            Total_Stock_Volume = 0
        ElseIf ws.Cells(i - 1, 1).Value <> Ticker Then
            ' Get the first day's opening price for the current ticker
            First_Open_Price = ws.Cells(i, 3).Value
        Else
            ' Add to Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i

    ' Print Greatest % Increase
    ws.Range("P2").Value = maxPercentIncreaseTicker
    ws.Range("Q2").Value = Format(maxPercentIncrease, "0.00%")

    ' Print Greatest % Decrease
    ws.Range("P3").Value = minPercentDecreaseTicker
    ws.Range("Q3").Value = Format(minPercentDecrease, "0.00%")

    ' Print Greatest Total Volume
    ws.Range("P4").Value = maxTotalVolumeTicker
    ws.Range("Q4").Value = maxTotalVolume
End Sub
