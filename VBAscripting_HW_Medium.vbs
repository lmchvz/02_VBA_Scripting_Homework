Sub sheets()

For Each ws In ActiveWorkbook.Worksheets
      ws.Activate
      Call Moderate

    Next ws

End Sub



Sub Moderate()
    Dim Ticker As Double
    Dim Stock_Volume As Double
    Dim cur_stock_sym As String
    Dim prev_stock_sym As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_pct_change As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastrow
        open_price = Cells(2, 3).Value
        cur_stock_sym = Cells(i, 1).Value
        day_stock_vol = Cells(i, 7).Value
        day_open_price = Cells(i, 3).Value
        day_close_price = Cells(i, 6).Value
        
        next_stock_sym = Cells(i + 1, 1).Value
        next_open_price = Cells(i + 1, 3).Value
        
        Stock_Volume = day_stock_vol + Stock_Volume
        If (open_price = 0) Then
            open_price = day_open_price
        End If
        
        If (cur_stock_sym <> next_stock_sym) Then
            close_price = day_close_price
            Cells(Ticker + 2, 9).Value = cur_stock_sym
            Cells(Ticker + 2, 12).Value = Stock_Volume
            Ticker = Ticker + 1
            
            yearly_change = (close_price - open_price)
            Cells(Ticker + 1, 10).Value = yearly_change
            If (close_price < open_price) Then
                Cells(Ticker + 1, 10).Interior.Color = vbRed
            Else
                Cells(Ticker + 1, 10).Interior.Color = vbGreen
            
            End If
            
            If (open_price = 0) Then
                Cells(Ticker + 1, 11).Value = "NA"
            Else
                Cells(Ticker + 1, 11).Value = yearly_change / open_price
            End If
          
            Cells(Ticker + 1, 11).NumberFormat = "0.00%"
            
            Stock_Volume = 0
            open_price = next_open_price
             
        End If
    
    Columns.AutoFit

    Next i


End Sub


