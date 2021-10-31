Sub stock_stats_formating():

Dim ticker As String

Dim ticker_amount As String

Dim last_row As Long

Dim price_open As Double

Dim price_close As Double

Dim yearly_change As Double

Dim percent_change As Double

Dim total_stock_volume As Double

Dim greatest_percent_increase As Double

Dim greatest_percent_increase_ticker As String

Dim greatest_percent_decrease As Double

Dim greatest_percent_decrease_ticker As String

Dim greatest_total_volume As Double

Dim greatest_total_volume_ticker As String

For Each ws In Worksheets

    ws.Activate
    
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ticker_amount = 0
    ticker = " "
    yearly_change = 0
    price_open = 0
    price_close = 0
    total_stock_volume = 0
    
    For i = 2 To last_row
    
        ticker = Cells(i, 1).Value
        
        If price_open = 0 Then
            price_open = Cells(i, 3).Value
        End If
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Code when ticker changes
        If Cells(i + 1, 1).Value <> ticker Then
        
            ticker_amount = ticker_amount + 1
            Cells(ticker_amount + 1, 9) = ticker
            
            price_close = Cells(i, 6)
            
            yearly_change = price_close - price_open
            
            Cells(ticker_amount + 1, 10).Value = yearly_change
            
            ' conditional formatting for yearly change
            If yearly_change > 0 Then
                Cells(ticker_amount + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(ticker_amount + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(ticker_amount + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            If price_open = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / price_open)
            End If
            
            
            Cells(ticker_amount + 1, 11).Value = Format(percent_change, "Percent")
            
            
            price_open = 0
            
            Cells(ticker_amount + 1, 12).Value = total_stock_volume
            
            total_stock = 0
        End If
        
    Next i
    
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        
        greatest_percent_increase = Cells(2, 11).Value
        greatest_percent_increase_ticker = Cells(2, 9).Value
        greatest_percent_decrease = Cells(2, 11).Value
        greatest_percent_decrease_ticker = Cells(2, 9).Value
        greatest_total_volume = Cells(2, 12).Value
        greatest_total_volume_ticker = Cells(2, 9).Value
        
        
        For i = 2 To last_row
        
            If Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = Cells(i, 11).Value
                greatest_percent_increase_ticker = Cells(i, 9).Value
            End If
            
            If Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = Cells(i, 11).Value
                greatest_percent_decrease_ticker = Cells(i, 9).Value
            End If
            
             If Cells(i, 11).Value > greatest_total_volume Then
                greatest_total_volume = Cells(i, 12).Value
                greatest_total_volume = Cells(i, 9).Value
            End If
        
        Next i
        
        
            Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
            Range("Q2").Value = Format(greatest_percent_increase, "Percent")
            Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
            Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
            Range("P4").Value = greatest_total_volume_ticker
            Range("Q4").Value = greatest_total_volume
            
Next ws
        

End Sub

