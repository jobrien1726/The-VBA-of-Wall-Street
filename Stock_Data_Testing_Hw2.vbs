Sub StockDataTesting()

Dim i As Long
Dim SummaryTableRow, SummaryTable2Row As Integer
Dim TotalStockVolume As Double
Dim lastrow, lastrow2 As Long
Dim WeekdayCount As Integer
Dim YearlyChange As Double
Dim PercentChange As Double

For Each ws In Worksheets

SummaryTableRow = 2
TotalStockVolume = 0
WeekdayCount = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        BegYrOpenPrice = ws.Cells((i - WeekdayCount), 3).Value
        YearlyChange = ws.Cells(i, 6).Value - BegYrOpenPrice
        
            If BegYrOpenPrice = 0 Then
                PercentChange = (YearlyChange / 1E-05) * 100
            Else
                PercentChange = (YearlyChange / BegYrOpenPrice) * 100
            End If
            
        ws.Range("I" & SummaryTableRow).Value = ws.Cells(i, 1).Value
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        
            If ws.Range("J" & SummaryTableRow).Value < 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            ElseIf ws.Range("J" & SummaryTableRow).Value > 0 Then
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            End If
            
        ws.Range("K" & SummaryTableRow).Value = PercentChange
        ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
        SummaryTableRow = SummaryTableRow + 1
        TotalStockVolume = 0
        WeekdayCount = 0
        
    Else
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        WeekdayCount = WeekdayCount + 1
        'MsgBox (WeekdayCount)
    End If
    
Next i

lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
SummaryTable2Row = 2

For i = 2 To lastrow2

If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & lastrow2)) Then
    ws.Range("P" & SummaryTable2Row).Value = ws.Cells(i, 9).Value
    ws.Range("Q" & SummaryTable2Row).Value = ws.Cells(i, 11).Value
    SummaryTable2Row = SummaryTable2Row + 1
End If
    
Next i

For i = 2 To lastrow2

If ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & lastrow2)) Then
    ws.Range("P" & SummaryTable2Row).Value = ws.Cells(i, 9).Value
    ws.Range("Q" & SummaryTable2Row).Value = ws.Cells(i, 11).Value
    SummaryTable2Row = SummaryTable2Row + 1
End If

Next i

For i = 2 To lastrow2

If ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & lastrow2)) Then
    ws.Range("P" & SummaryTable2Row).Value = ws.Cells(i, 9).Value
    ws.Range("Q" & SummaryTable2Row).Value = ws.Cells(i, 12).Value
End If

Next i

ws.Columns("A:Q").AutoFit

Next ws

End Sub
