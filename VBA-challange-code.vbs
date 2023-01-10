Sub DataExtraction()
For Each ws In Worksheets
Dim column As Integer
Dim rowkeep As Integer
Dim counter1 As Integer
Dim counter2 As Integer
Dim YearlyChange As Double
Dim TotalStockVolume As LongLong
column = 1
counter1 = 0
counter2 = 1
rowkeep = 2
TotalStockVolume = 0
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
YearlyChange = ws.Cells(i, 6).Value - ws.Cells(i - counter1, 3).Value
ws.Cells(rowkeep, 9).Value = ws.Cells(i, 1).Value
ws.Cells(rowkeep, 10).Value = YearlyChange
ws.Cells(rowkeep, 11).Value = YearlyChange / ws.Cells(i - counter1, 3).Value
ws.Cells(rowkeep, 12).Value = TotalStockVolume
rowkeep = rowkeep + 1
counter1 = 0
counter2 = counter2 + 1
TotalStockVolume = 0
Else: counter1 = counter1 + 1
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
End If
Next i
For i = 2 To counter2
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
End If
ws.Cells(i, 11).NumberFormat = "0.00%"
Next i
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(counter2, 11)))
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(counter2, 11)))
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(counter2, 12)))
For i = 2 To counter2
If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
End If
Next i
Next ws
End Sub
