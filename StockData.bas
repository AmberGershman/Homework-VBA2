Attribute VB_Name = "Module1"
Sub alphatest()
    Dim ticker As String
    Dim ticker_total As Double
    ticker_total = 0
    Dim ticker_table_row As String
    ticker_table_row = 2
    stock_open = Cells(2, 3).Value
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("K2:K" & lastrow).NumberFormat = "0.00%"
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To lastrow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            ticker_total = ticker_total + Cells(i, 7).Value
            stock_close = Cells(i, 6).Value
            yearly_change = stock_close - stock_open
            per_change = yearly_change / stock_open
            Range("I" & ticker_table_row).Value = ticker
            Range("J" & ticker_table_row).Value = yearly_change
                If yearly_change > 0 Then
                    Range("J" & ticker_table_row).Interior.ColorIndex = 43
                Else
                    Range("J" & ticker_table_row).Interior.ColorIndex = 46
                End If
            Range("K" & ticker_table_row).Value = per_change
            Range("L" & ticker_table_row).Value = ticker_total
            ticker_table_row = ticker_table_row + 1
            ticker_total = 0
            yearly_change = 0
            per_change = 0
        Else
            ticker_total = ticker_total + Cells(i, 7).Value
        End If
    Next i
    
            
End Sub

Sub Bonus()

    Dim tickerrange As Range
    Set tickerrange = Range("K2: K3010")
    Dim stocktotals As Range
    Set stocktotals = Range("L2: L3010")
    
    Range("P2, P3").NumberFormat = "0.00%"
    Dim total_stock As Double
    
    Cells(2, 14).Value = "Greatest Percent Increase"
    Cells(3, 14).Value = "Greatest Percent Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Range("I1:P1").Font.Bold = True
    Range("I1:P1").HorizontalAlignment = xlCenter
    Range("N2:N4").Font.Bold = True
    Range("N2:N4").HorizontalAlignment = xlRight
    Range("O2:O3").HorizontalAlignment = xlCenter
    
    max_stock = Application.WorksheetFunction.Max(tickerrange)
    Range("P2").Value = max_stock
    min_stock = Application.WorksheetFunction.Min(tickerrange)
    Range("P3").Value = min_stock
    total_stock = Application.Sum(stocktotals)
    Range("P4").Value = total_stock
    
    For i = 2 To 3010
        If Cells(i, 11).Value = max_stock Then
            Range("O2").Value = Cells(i, 9)
        End If
        If Cells(i, 11).Value = min_stock Then
            Range("O3").Value = Cells(i, 9)
        End If
        
    Next i
    
    
End Sub
