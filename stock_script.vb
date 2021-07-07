Sub stock_script()
    Dim ws As Worksheet
    
    'Looping through all worksheets
    For Each ws In Worksheets
        ws.Activate
        
        'Variable decalaration
        Dim ticker_r As String
        Dim volume_r, sum_vol, lr, r As Long
        Dim open_r, close_r As Double
        
        'Assigning headers
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change ($)"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        
        sum_vol = 0
        
        'Last Used Row
        lr = Cells(Rows.Count, 1).End(xlUp).Row
        
        out_r = 2
        open_r = Cells(2, 3)
    
        For r = 2 To lr
            'Current values of ticker, closing price and volume
            ticker_r = Cells(r, 1)
            close_r = Cells(r, 6)
            volume_r = Cells(r, 7)
            
            'Finding total volume
            sum_vol = sum_vol + volume_r
            
            If ticker_r <> Cells(r + 1, 1) Then
                Cells(out_r, 9) = ticker_r
                Cells(out_r, 10) = close_r - open_r
                
                'Accounting for stock prices with opening price =0
                If open_r <> 0 Then
                    Cells(out_r, 11) = (close_r - open_r) / open_r
                End If
                
                Cells(out_r, 12) = sum_vol
                
                open_r = Cells(r + 1, 3)
                sum_vol = 0
                
                'Assigning color formatting
                If Cells(out_r, 10) < 0 Then
                    Cells(out_r, 10).Interior.Color = vbRed
                ElseIf Cells(out_r, 10) >= 0 Then
                    Cells(out_r, 10).Interior.Color = vbGreen
                End If
                
                out_r = out_r + 1
                
            End If
              
        Next r
        Range("K:K").NumberFormat = "0.00%"
    
        
        'Bonus
            
        Dim i, new_last As Long
        Dim min_percent, max_percent, max_vol As Double
        Dim min_tick, max_tick, max_vol_tick As String
        
        new_last = Cells(Rows.Count, 9).End(xlUp).Row
    
        
        Range("O2") = "Greatest  % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        'Finding the max % increase, max % decrease and max volume
        min_percent = 0
        max_percent = 0
        max_vol = 0
        
        For i = 2 To new_last
            If Cells(i, 11) < min_percent Then
                min_percent = Cells(i, 11)
                min_tick = Cells(i, 9)
                
            End If
            If Cells(i, 11) > max_percent Then
                max_percent = Cells(i, 11)
                max_tick = Cells(i, 9)
                
            End If
            If Cells(i, 12) > max_vol Then
                max_vol = Cells(i, 12)
                max_vol_tick = Cells(i, 9)
                
            End If
                
        Next i
        
        Range("P2") = max_tick
        Range("Q2") = max_percent
        
        Range("P3") = min_tick
        Range("Q3") = min_percent
        
        Range("P4") = max_vol_tick
        Range("Q4") = max_vol
        
        Range("Q2:Q3").NumberFormat = "0.00%"
        Range("Q4").NumberFormat = "0.00E+00"
    Next ws
End Sub

