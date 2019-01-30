Sub tickerTableIII()

Dim i, ti As Integer
Dim total, openingPrice, closingPrice, gtInc, gtDec, gtVol As Double
Dim ticker, gtIncTicker, gtDecTicker, gtVolTicker As String

For Each ws In Worksheets

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ws.Cells(2, 1).Value
    openingPrice = ws.Cells(2, 3).Value
    total = ws.Cells(2, 7).Value
    gtInc = 0
    gtDec = 0
    gtVol = 0

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Volume"

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
           If gtInc < ws.Cells(ti, 11).Value Then
                gtInc = ws.Cells(ti, 11).Value
                gtIncTicker = ticker
            End If
            If gtDec > ws.Cells(ti, 11).Value Then
                gtDec = ws.Cells(ti, 11).Value
                gtDecTicker = ticker
            End If
            If gtVol < total Then
                gtVol = total
                gtVolTicker = ticker
            End If

            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            total = ws.Cells(i, 7).Value
            ti = ti + 1
        End If
    Next i

    ws.Cells(2, 16).Value = gtIncTicker
    ws.Cells(2, 17).Value = gtInc
    ws.Cells(2, 17).Style = "Percent"
    ws.Cells(3, 16).Value = gtDecTicker
    ws.Cells(3, 17).Value = gtDec
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(4, 16).Value = gtVolTicker
    ws.Cells(4, 17).Value = gtVol
    ws.Columns("A:Q").AutoFit


Next ws

End Sub
