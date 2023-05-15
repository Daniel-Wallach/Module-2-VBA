Attribute VB_Name = "Module1"
Sub StockDataYearlyChange()

    'Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim vol As Double
    Dim counter As Long
    Dim maxIncr As Double
    Dim maxDecr As Double
    Dim maxVol As Double
    Dim maxIncrTicker As String
    Dim maxDecrTicker As String
    Dim maxVolTicker As String
    
    For Each ws In Worksheets
        
        'Assign values to variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ticker = ""
        opening = 0
        closing = 0
        yearlyChange = 0
        percentChange = 0
        vol = 0
        counter = 2
        maxIncr = ws.Cells(2, 11).Value
        maxDecr = ws.Cells(2, 11).Value
        maxVol = 0
        maxIncrTicker = ""
        maxDecrTicker = ""
        maxVolTicker = ""
        
        'Create headers
        ws.Cells(1, 9).Value = "Tickers"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 11).NumberFormat = "0.00%"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Loop through each row
        For i = 2 To lastRow
            
            'If next row is new ticker symbol
            If ws.Cells(i, 1).Value <> ticker Then
                
                'Output the last ticker's results if not first ticker
                If Not (ticker = "") Then
                
                    ws.Cells(counter, 9).Value = ticker
                    ws.Cells(counter, 10).Value = yearlyChange
                    ws.Cells(counter, 11).Value = percentChange
                    ws.Cells(counter, 11).NumberFormat = "0.00%"
                    ws.Cells(counter, 12).Value = vol
                    
                    'Move to next row
                    counter = counter + 1
                    
                End If
                
                'Reset variables for new ticker
                ticker = ws.Cells(i, 1).Value
                opening = ws.Cells(i, 3).Value
                yearlyChange = 0
                percentChange = 0
                vol = 0
                
            End If
            
            'Calculate yearly and percentage change
            closing = ws.Cells(i, 6).Value
            yearlyChange = closing - opening
            percentChange = yearlyChange / opening
            vol = vol + ws.Cells(i, 7).Value
            
            'Output last ticker's results
            If i = lastRow Then
                ws.Cells(counter, 9).Value = ticker
                ws.Cells(counter, 10).Value = yearlyChange
                ws.Cells(counter, 11).Value = percentChange
                ws.Cells(counter, 11).NumberFormat = "0.00%"
                ws.Cells(counter, 12).Value = vol
                
            End If
            
            'Highlight positive change in green and negative change in red
            If yearlyChange > 0 Then
                ws.Range("J" & counter).Interior.ColorIndex = 4
            ElseIf yearlyChange < 0 Then
                ws.Range("J" & counter).Interior.ColorIndex = 3
            Else
                ws.Range("J" & counter).Interior.ColorIndex = 0
                    
            End If
            
            'Check for greatest % increase, % decrease, and total volume
            If ws.Cells(i, 11).Value > maxIncr Then
                maxIncr = ws.Cells(i, 11).Value
                maxIncrTicker = ws.Cells(i, 9)
            ElseIf ws.Cells(i, 11).Value < maxDecr Then
                maxDecr = ws.Cells(i, 11).Value
                maxDecrTicker = ws.Cells(i, 9)
                
            End If
            
            If vol > maxVol Then
                maxVol = vol
                maxVolTicker = ticker
                
            End If
            
        Next i
        
        'Output greatest % increase, % decrease, and total volume
        ws.Cells(2, 15).Value = maxIncrTicker
        ws.Cells(2, 16).Value = maxIncr
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = maxDecrTicker
        ws.Cells(3, 16).Value = maxDecr
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = maxVolTicker
        ws.Cells(4, 16).Value = maxVol
        
    Next ws
    
End Sub
