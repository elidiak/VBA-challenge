Attribute VB_Name = "Module1"
Sub StockSummary()
    'Define the worksheet variable
    Dim ws As Worksheet
    
    'Worksheet loop this is what copilot gave me
    For Each ws In ThisWorkbook.Sheets
    
        'Define variables for the loop
        Dim StockTicker As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim QuarterlyChange As Double
        Dim TotalVolume As Double
        Dim SummaryRow As Double
    
        'Set the starting value for each variable
        StockTicker = ""
        OpeningPrice = 0
        ClosingPrice = 0
        QuarterlyChange = 0
        TotalVolume = 0
        SummaryRow = 2
    
        'Add the Summary Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'Add the headers and row lables for Greatest percent increase, decrease, and greatest total volume
    
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Volume"
    
        'Build the loop that tallies the summary
    
        Dim MyRow As Double
        Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        For MyRow = 2 To LastRow
    
            If ws.Cells(MyRow - 1, 1).Value <> ws.Cells(MyRow, 1).Value Then
            'Read in the ticker and opening price when the current row has a new ticker
        
            'Grab the new ticker
        
            StockTicker = ws.Cells(MyRow, 1).Value
            ws.Cells(SummaryRow, 9) = StockTicker
        
            'Grab the opening price
            OpenPrice = ws.Cells(MyRow, 3)
            
            'Grab Total Volume
            TotalVolume = TotalVolume + ws.Cells(MyRow, 7).Value
    
            ElseIf ws.Cells(MyRow + 1, 1).Value <> ws.Cells(MyRow, 1).Value Then
            
            'Grab the Closing Price
            ClosePrice = ws.Cells(MyRow, 6).Value
        
            'Calculate the quarterly change
            QuarterlyChange = ClosePrice - OpenPrice
            ws.Cells(SummaryRow, 10).Value = QuarterlyChange
        
            'Format the Quarterly Value
            If QuarterlyChange < 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            ElseIf QuarterlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            End If
        
            'Calculate the percentage change
            ws.Cells(SummaryRow, 11) = QuarterlyChange / OpenPrice
        
            'Set the cell formatting
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
        
            'Add the current row's closing price and volume
            TotalVolume = TotalVolume + ws.Cells(MyRow, 7).Value
            
            'Write the total volume to the summary
            ws.Cells(SummaryRow, 12).Value = TotalVolume

            'Increase the summary row so I don't overwrite
            SummaryRow = SummaryRow + 1

            
            'Reset the variables
            QuarterlyChange = 0
            OpeningPrice = 0
            ClosingPrice = 0
            TotalVolume = 0

            Else
            
                'Add the current row's closing price and volume
                TotalVolume = TotalVolume + ws.Cells(MyRow, 7).Value
            
            End If

        Next MyRow
    
        'We need to create the second summary
        'For percentage increase we need a loop
        Dim StockTicker2 As String
        Dim StockTicker3 As String
        Dim StockTicker4 As String
        Dim PercentageIncrease As Double
        Dim PercentageDecrease As Double
        Dim TotalVolume2 As Double
        Dim SummaryRow2 As Double
    
        'Set the starting value for each
        StockTicker2 = ""
        StockTicker3 = ""
        StockTicker4 = ""
        PercentageIncrease = 0
        PercentageDecrease = 0
        TotalVolume2 = 0
        SummaryRow2 = 2
    
        'Loop variables
        Dim MyRow2 As Double
        Dim LastRow2 As Double
        LastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

        'Core loop for Second Summary
        For MyRow2 = 2 To LastRow2
        If ws.Cells(MyRow2, 11).Value > PercentageIncrease Then
            'Copy the ticker over
            StockTicker2 = ws.Cells(MyRow2, 9).Value
            PercentageIncrease = ws.Cells(MyRow2, 11).Value
        End If

        'If statement for Percentage Decrease
        If ws.Cells(MyRow2, 11).Value < PercentageDecrease Then
            'Copy the ticker over
            StockTicker3 = ws.Cells(MyRow2, 9).Value
            PercentageDecrease = ws.Cells(MyRow2, 11).Value
        End If

        'If statement for Greatest Total Volume2
        If ws.Cells(MyRow2, 12).Value > TotalVolume2 Then
            'Copy the ticker over
            StockTicker4 = ws.Cells(MyRow2, 9).Value
            TotalVolume2 = ws.Cells(MyRow2, 11).Value
        End If
        Next MyRow2
    
        'Push Summary 2 values
        'Percentage Increase
    
        ws.Cells(2, 15).Value = StockTicker2
        ws.Cells(2, 16).Value = PercentageIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
    
        'Percentage Decrease
        ws.Cells(3, 15).Value = StockTicker3
        ws.Cells(3, 16).Value = PercentageDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
    
        'Highest Volume
        ws.Cells(4, 15).Value = StockTicker4
        ws.Cells(4, 16).Value = TotalVolume2
        
    Next ws
End Sub
Sub ResetSummary()
    'From Copilot, clear the last try
        'Define the worksheet variable
    Dim ws As Worksheet
    
    'Worksheet loop this is what copilot gave me
    For Each ws In ThisWorkbook.Sheets
        ws.Range("I:P").ClearContents
        ws.Range("I:P").ClearFormats
    Next ws
    
End Sub

