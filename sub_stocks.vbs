Sub stocks():
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Total Stock Volume"
    Dim yearlyChange As Long
    Dim loc As Long
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        yearlyChange = yearlyChange + (Cells(i, 6).Value - Cells(i, 3).Value)
        If Cells(i, 1) <> Cells(i + 1, 1) Then
                loc = loc + 1
                'New Ticker
                Cells(loc, 9) = Cells(loc, 1)
                Cells(loc, 10).Value = yearlyChange
    Next i

End Sub
