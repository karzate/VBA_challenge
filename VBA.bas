Attribute VB_Name = "Module1"
Sub Analysis():

'declare variables
Dim ticker As String
Dim number_tickers As Integer
Dim closing_price, opening_price, percent_change, volume As Double
Dim lastRowState As Long
Dim yearly_change, total_stock_volume, greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease, greatest_stock_volume As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume_ticker As String

' loop ws in the workbook
For Each ws In Worksheets

    ' worksheet active
    ws.Activate

    ' last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns for each worksheet
    ws.Range("l1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    ' variables for worksheets
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Skip the HR
    For x = 2 To lastRowState

        ' value ticker symbol
        ticker = Cells(x, 1).Value
        
        ' opening price
        If opening_price = 0 Then
            opening_price = Cells(x, 3).Value
        End If
        
        ' total stock volume values
        total_stock_volume = total_stock_volume + Cells(x, 7).Value
    
        If Cells(x + 1, 1).Value <> ticker Then
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' end of closing price for ticker
            closing_price = Cells(x, 6)
            
            ' yearly change value
            yearly_change = closing_price - opening_price
            
            ' yearly change value to each worksheet.
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' shade green
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' shade red
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' shade yellow
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' percent change
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Format percent_change
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' opening price
            opening_price = 0
            
            ' total stock volume in each worksheet
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' total stock volume 0
            total_stock_volume = 0
        End If
        
    Next x
    
    ' greatest percent increase, greatest percent decrease, and greatest total volume
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    '  variables and values
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' skip the HR
    For i = 2 To lastRowState
    
        ' greatest percent increase
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        '  greatest percent decrease
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' greatest stock volume
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Values greatest percent increase, decrease, and stock volume
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub

