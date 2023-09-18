Attribute VB_Name = "Module1"
Sub alphabetical_testing()

    Dim ws As Worksheet

    'Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        'Assuming headers are in the first row, start from the second row
        Dim startRow As Long: startRow = 2
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        'Variables for holding the yearly change, percentage change, total volume, etc.
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        Dim percChange As Double
        Dim totalVolume As Double: totalVolume = 0

        Dim greatestPercIncrease As Double: greatestPercIncrease = 0
        Dim greatestPercDecrease As Double: greatestPercDecrease = 0
        Dim greatestTotalVolume As Double: greatestTotalVolume = 0

        'Loop through each row in the current worksheet
        For i = startRow To lastRow
            'If we are at the first row or if the ticker symbol changes
            If i = startRow Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                openPrice = ws.Cells(i, 3).Value
            End If

            'Accumulate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            'If the ticker symbol changes or if we are at the last row
            If i = lastRow Or ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percChange = yearlyChange / openPrice
                Else
                    percChange = 0
                End If

                'Output the ticker symbol, yearly change, percentage change, and total volume
                Dim summaryRow As Long: summaryRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row + 1
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percChange
                ws.Cells(summaryRow, 12).Value = totalVolume

                'Reset values for the next stock
                totalVolume = 0

                'Check for greatest percentage increase, decrease, and total volume
                If percChange > greatestPercIncrease Then greatestPercIncrease = percChange
                If percChange < greatestPercDecrease Then greatestPercDecrease = percChange
                If totalVolume > greatestTotalVolume Then greatestTotalVolume = totalVolume
            End If
        Next i

        'Output greatest percentage increase, decrease, and total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = greatestPercIncrease
        ws.Cells(3, 16).Value = greatestPercDecrease
        ws.Cells(4, 16).Value = greatestTotalVolume

        'Conditional formatting
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).FormatConditions(1).Interior.Color = RGB(0, 255, 0) 'Green
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).FormatConditions(2).Interior.Color = RGB(255, 0, 0) 'Red

    Next ws

End Sub


