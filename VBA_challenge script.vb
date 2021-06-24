Sub VBA_challenge()

For Each ws In Worksheets

    'Set variable for last row in each worksheet
    Dim LastRow As LongLong
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Set variable for holding the ticker
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Percent_Change As Double
    Dim Volume_Total As Double
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Volume_Total = 0

    'Add headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    For i = 2 To LastRow

'Determine Year Open value and location for each ticker
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        Year_Open = ws.Cells(i, 3).Value

            'Omit calculation of percent change when denominator is "0'"
            ElseIf Year_Open = 0 Then
            Percent_Change = 0

            'Extract each unique ticker and Year Close
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Each unique ticker will come from...
                Ticker = ws.Cells(i, 1).Value

                'Year Close will ccome from...
                Year_Close = ws.Cells(i, 6).Value

                Yearly_Change = Year_Close - Year_Open

                Percent_Change = Yearly_Change / Year_Open

                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                ' Print the Ticker, Yearly Change, Percent Change and Volume Total in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                ws.Range("L" & Summary_Table_Row).Value = Volume_Total

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset Volume_Total
                Volume_Total = 0

            Else

            'Add to the Volume_Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value

        End If

    Next i

'Calculate Highest Percent Increase and assign it to appropriate cell
ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row))
ws.Cells(2, 17).NumberFormat = "0.00%"
Greatest_Percent_Increase = ws.Cells(2, 17).Value

'Calculate Highest Percent Decrease and assign it to appropriate cell
ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row))
ws.Cells(3, 17).NumberFormat = "0.00%"
Greatest_Percent_Decrease = ws.Cells(3, 17).Value

'Calculate Highest Total Volume and assign it to appropriate cell
ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row))
Greatest_Total_Volume = ws.Cells(4, 17).Value

'Create another For Loop for color coding Yearly Change and to match Highest Percent Increase to corresponing ticker
    For j = 2 To LastRow

        'Turn cell to red when yearly change is negative
        If ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3

        'Turn cell to green when yearly change is positive
        ElseIf ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4

        End If
        
        'Grab corresponding ticker for the greatest percent increase and assign it to appropriate cell
        If ws.Cells(j, 11).Value = Greatest_Percent_Increase Then
        ws.Cells(2, 16).Value = ws.Cells(j, 9).Value

        ''Grab corresponding ticker for the greatest percent decrease and assign it to appropriate cell
        ElseIf ws.Cells(j, 11).Value = Greatest_Percent_Decrease Then
        ws.Cells(3, 16).Value = ws.Cells(j, 9).Value

        'Grab corresponding ticker for the greatest total volume and assign it to appropriate cell
        ElseIf ws.Cells(j, 12).Value = Greatest_Total_Volume Then
        ws.Cells(4, 16).Value = ws.Cells(j, 9).Value

        End If

    Next j

'Autofit all columns
ws.UsedRange.Columns.AutoFit

'Chnage formating to scientific notation for small summary table 
ws.Range("Q4").NumberFormat = "0.0000E+00"

Next ws

End Sub