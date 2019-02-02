Sub stock():

    Dim stock_name As String
    Dim stock_total As Double
    stock_total = 0

    Dim results_table As Long
    results_table = 2

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Value"
    
    For i = 2 To lastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            stock_name = Cells(i, 1).Value
            stock_total = stock_total + Cells(i, 7).Value

            Range("I" & results_table).Value = stock_name

            Range("J" & results_table).Value = stock_total

            results_table = results_table + 1
        
            stock_total = 0
        
        Else

            stock_total = stock_total + Cells(i, 7).Value

        End If

    Next i

End Sub


