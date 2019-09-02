Attribute VB_Name = "Module1"
Sub Stock_Analysis()

For Each ws In Worksheets

Dim tickers(10000), GI_ticker, GD_ticker, GV_ticker As String
Dim yearopenings(10000), yearclosings(10000), yearvolume(10000), ytdvariance(10000), greater_decrease, greater_increase, greater_volume As Double
Dim i, nrows, counter As Long

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

nrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
counter = -1
greater_decrease = 0
greater_increase = 0
greater_volume = 0

For i = 2 To nrows + 1
    If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
        counter = counter + 1
        tickers(counter) = ws.Cells(i, 1)
        yearopenings(counter) = ws.Cells(i, 3)
        yearclosings(counter) = ws.Cells(i, 6)
        yearvolume(counter) = ws.Cells(i, 7)
        If counter > 0 Then
            ws.Range("I" & (counter + 1)).Value = tickers(counter - 1)
            ws.Range("J" & (counter + 1)).Value = yearclosings(counter - 1) - yearopenings(counter - 1)
            ws.Range("K" & (counter + 1)).Value = ytdvariance(counter - 1)
            ws.Range("L" & (counter + 1)).Value = yearvolume(counter - 1)
            
            If ws.Range("J" & (counter + 1)).Value > 0 Then
                ws.Range("J" & (counter + 1)).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & (counter + 1)).Value < 0 Then
                ws.Range("J" & (counter + 1)).Interior.ColorIndex = 3
            End If
            If greater_increase < ytdvariance(counter - 1) Then
                greater_increase = ytdvariance(counter - 1)
                GI_ticker = tickers(counter - 1)
            End If
            If greater_decrease > ytdvariance(counter - 1) Then
                greater_decrease = ytdvariance(counter - 1)
                GD_ticker = tickers(counter - 1)
            End If
            If greater_volume < yearvolume(counter - 1) Then
                greater_volume = yearvolume(counter - 1)
                GV_ticker = tickers(counter - 1)
            End If

        End If
    Else
        yearclosings(counter) = ws.Cells(i, 6)
        yearvolume(counter) = yearvolume(counter) + ws.Cells(i, 7)
        If yearopenings(counter) <> 0 Then
            ytdvariance(counter) = (yearclosings(counter) - yearopenings(counter)) / yearopenings(counter)
        End If
    End If
Next i

ws.Range("N2").Value = "Greater % Increase"
ws.Range("N3").Value = "Greater % Decrease"
ws.Range("N4").Value = "Greater Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

ws.Range("O2").Value = GI_ticker
ws.Range("O3").Value = GD_ticker
ws.Range("O4").Value = GV_ticker

ws.Range("P2").Value = greater_increase
ws.Range("P3").Value = greater_decrease
ws.Range("P4").Value = greater_volume

ws.Columns("A:P").EntireColumn.AutoFit
ws.Columns("K:K").NumberFormat = "0.0%"
ws.Columns("L:L").NumberFormat = "#,##0.0,,""m"""

ws.Range("P2:P3").NumberFormat = "0.0%"
ws.Range("P4").NumberFormat = "#,##0.0,,""m"""
Next ws

End Sub
