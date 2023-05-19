Attribute VB_Name = "Module2"
Sub StockLoop()

    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim Strt_opn As Double
        Dim TickerName As String
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        
        Dim TotalStockVolume As Double
        TotalStockVolume = 0

        Dim SummaryTableRow As Integer
        SummaryTableRow = 2

        lastrow = ws.Range("A1").End(xlDown).Row

        Strt_opn = 2


        For i = 2 To lastrow

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                TickerName = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(Strt_opn, 3).Value
                PercentageChange = YearlyChange / ws.Cells(Strt_opn, 3).Value
                
                Strt_opn = i + 1

                ws.Range("I" & SummaryTableRow).Value = TickerName
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentageChange

                TotalStockVolume = 0
                SummaryTableRow = SummaryTableRow + 1
        

            Else

                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            End If

        Next i
        increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        ws.Range("P2") = ws.Cells(increase_index + 1, 9)
        
        decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        ws.Range("P3") = ws.Cells(decrease_index + 1, 9)
        
        increase_indexTotal = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
        ws.Range("P4") = ws.Cells(increase_indexTotal + 1, 9)
        
        ws.Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
        ws.Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
        ws.Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & lastrow)) * 100
        
    Next ws
End Sub


