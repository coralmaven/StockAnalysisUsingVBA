Sub tickerTableI()

Dim i, total, ti As Integer
Dim ticker As String


For Each ws In Worksheets

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ws.Cells(2, 1).Value
    total = ws.Cells(2, 7).Value
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ti = 2
    
    For i = 3 To lastRow + 1
        If ticker = ws.Cells(i, 1) Then
            total = total + ws.Cells(i, 7).Value
        Else
            ws.Cells(ti, 9).Value = ticker
            ws.Cells(ti, 10).Value = total
            ticker = ws.Cells(i, 1).Value
            total = ws.Cells(i, 7).Value
            ti = ti + 1
        End If
    Next i

Next ws

End Sub