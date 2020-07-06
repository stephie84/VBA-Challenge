Sub Stocks()


'Loop through all Worksheets
For Each ws In Worksheets


'Column Headings for Worksheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

'Variables
Dim LastRow As Long
Dim i As Long

Dim TickerName As String

Dim StockVolume As Double


Dim SummaryTable As Integer
SummaryTable = 2



'Define LastRow
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop for ticker, total stock volume

For i = 2 To LastRow
    'Check if Ticker has changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
    'Ticker Name
    TickerName = ws.Cells(i, 1).Value
    
    'Stock Volume
    StockVolume = StockVolume + ws.Cells(i, 7).Value
    
    'Print Ticker
    ws.Range("I" & SummaryTable).Value = TickerName
    
    'Print Total Stock Volume
    ws.Range("J" & SummaryTable).Value = StockVolume
    
    
    SummaryTable = SummaryTable + 1
    
    StockVolume = 0
    
    
    'If Ticker has not changed
    Else
    
    StockVolume = StockVolume + ws.Cells(i, 7).Value
   
       
    
   End If
    
Next i
Next ws

End Sub

