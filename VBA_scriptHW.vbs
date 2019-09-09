Sub loopforticker()

For Each ws In Worksheets

Dim tickername As String
Dim tickervolume As Double
tickervolume = 0
Dim lastrow As Long

Dim results_tickers As Long
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim open_price As Double
open_price = ws.Cells(2, 3).Value
Dim close_price As Double

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

results_tickers = 2

For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        tickername = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        ws.Range("I" & results_tickers).Value = tickername
        ws.Range("L" & results_tickers).Value = total_volume
        close_price = ws.Cells(i, 6).Value
        yearly_change = (close_price - open_price)
        ws.Range("J" & results_tickers).Value = yearly_change

        
        
        If open_price = 0 Then
          percent_change = 0
          
        Else
            
            percent_change = (yearly_change / open_price)
        End If
        
        ws.Range("K" & results_tickers).Value = percent_change
        ws.Range("K" & results_tickers).NumberFormat = "0.00%"
        results_tickers = results_tickers + 1
        total_volume = 0
        open_price = ws.Cells(i + 1, 3)
    Else
        total_volume = total_volume + ws.Cells(i, 7).Value
        End If
Next i




lastrow_results = ws.Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrow_results
    If ws.Cells(i, 10) < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
    End If
Next i

Next ws

End Sub

