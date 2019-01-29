Sub tickerTableII()

Dim i, ti As Integer
Dim total, openingPrice, closingPrice As Double
Dim ticker As String

For Each ws In Worksheets

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ws.Cells(2, 1).Value
    openingPrice = ws.Cells(2, 3).Value
    total = ws.Cells(2, 7).Value
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ti = 2
    
    For i = 3 To lastRow + 1
    
        If ticker = ws.Cells(i, 1) Then
            total = total + ws.Cells(i, 7).Value
            closingPrice = ws.Cells(i, 6).Value
        Else
            ws.Cells(ti, 9).Value = ticker
            
            ws.Cells(ti, 10).Value = openingPrice - closingPrice
            
            If openingPrice = 0 Then
                openingPrice = 0.0001
            End If
            
            ws.Cells(ti, 11).Value = (openingPrice - closingPrice) / openingPrice
            If ws.Cells(ti, 11).Value >= 0 Then
                ws.Cells(ti, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(ti, 11).Interior.ColorIndex = 3
            End If
            ws.Cells(ti, 11).Style = "Percent"
            ws.Cells(ti, 12).Value = total
            
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            total = ws.Cells(i, 7).Value
            ti = ti + 1
        End If
        
    Next i
    ws.Columns("A:L").AutoFit
Next ws

End Sub

