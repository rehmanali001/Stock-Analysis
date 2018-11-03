Sub StockAnalysis():
    
    For Each ws In Worksheets
    
    Dim Ticker As String
    Dim Volume As Double
    Dim Row As Integer
    
    Volume = 0
    Row = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Volume = Volume + Cells(i, 7).Value
            ws.Range("I" & Row).Value = Ticker
            ws.Range("I1").Value = "Ticker"
            ws.Range("J" & Row).Value = Volume
            ws.Range("J1").Value = "Total Stock Volume"
                
            Row = Row + 1
            Volume = 0
        Else
            Volume = Volume + ws.Cells(i, 7).Value
        End If
    Next i
    Next
End Sub

