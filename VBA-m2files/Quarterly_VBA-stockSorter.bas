Attribute VB_Name = "Module1"

Sub stockSorter()

Dim ws As Worksheet
Dim tickerName As String
Dim infoPlacer As Integer
Dim stockVolume As Variant
Dim lastRow As Variant
Dim greatInc As Double
Dim greatDec As Double
Dim holdInc As Double
Dim holdDec As Double
Dim tickerNameI As String
Dim tickerNameD As String
Dim tickerNameG As String
Dim greatVolume As Variant


Dim columnCheck As Integer
Dim closePrice As Double
Dim openPrice As Double
Dim quarterlyChange As Double
Dim percentChange As Double
Dim stockSet As Long
Dim i As Long

For Each ws In Worksheets




ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N1").Value = "Greatest % Increase"
ws.Range("N2").Value = "Greatest % Decrease"
ws.Range("N3").Value = "Greatest Total Volume"


infoPlacer = 2
stockVolume = 0

closePrice = 0
openPrice = 0
quarterlyChange = 0
stockSet = 0
percentChange = 0
holdInc = 0
holdDec = 0
greatVolume = 0


lastRow = Cells(Rows.Count, 1).End(xlUp).Row


    For i = 2 To lastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        closePrice = ws.Cells(i, 6)
        openPrice = ws.Cells(i - stockSet, 3)
        quarterlyChange = closePrice - openPrice
        percentChange = (quarterlyChange / openPrice)
        
        
        ws.Cells(infoPlacer, 9).Value = tickerName
        ws.Cells(infoPlacer, 10).Value = quarterlyChange
       
        ws.Cells(infoPlacer, 11).Value = FormatPercent(percentChange)
        ws.Cells(infoPlacer, 12).Value = stockVolume
        
        
        If percentChange < 0 Then
        ws.Cells(infoPlacer, 10).Interior.ColorIndex = 3
        
            Else
            ws.Cells(infoPlacer, 10).Interior.ColorIndex = 4
        
        End If
        
        stockVolume = 0
        infoPlacer = infoPlacer + 1
        stockSet = 0
        
        Else
        stockVolume = stockVolume + ws.Cells(i, 7).Value
        tickerName = ws.Cells(i, 1).Value
        stockSet = stockSet + 1
          
        End If
        
    If ws.Cells(i, 11).Value > holdInc Then
        holdInc = ws.Cells(i, 11).Value
        tickerNameI = ws.Cells(i, 9)
        
        ElseIf ws.Cells(i, 11).Value < holdDec Then
        holdDec = ws.Cells(i, 11).Value
        tickerNameD = ws.Cells(i, 9)
        
        Else
        
        End If
        
        
    If ws.Cells(i, 12).Value > greatVolume Then
        greatVolume = ws.Cells(i, 12).Value
        tickerNameG = ws.Cells(i, 9).Value
        Else
        End If
        
    
    
    Next i
    
ws.Range("O1").Value = tickerNameI
ws.Range("O2").Value = tickerNameD
ws.Range("O3").Value = tickerNameG


ws.Range("P1").Value = FormatPercent(holdInc)
ws.Range("P2").Value = FormatPercent(holdDec)
ws.Range("P3").Value = greatVolume

Next ws

        
End Sub

