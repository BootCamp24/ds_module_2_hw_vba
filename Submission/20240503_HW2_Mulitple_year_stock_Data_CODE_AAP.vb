
Sub stockAnalysis():


    For Each ws In ThisWorkbook.Worksheets
    
        ' need to define ticker, open_price, close_price, vol_total, i and j
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim vol_total As Double
        Dim row_count As Double
        Dim delta_price As Double
        Dim percent_change As Double
        Dim i As Double
        Dim j As Integer
        ' count number of row that needs to be evaluated
        row_count = ws.Application.WorksheetFunction.CountIf(ws.Range("a:a"), "<>")
        j = 2
        open_price = 0
        open_price = ws.Cells(2, 3).Value
        ticker = ws.Cells(2, 1).Value
        For i = 2 To row_count
            If (ticker = ws.Cells(i + 1, 1).Value) Then
                vol_total = vol_total + ws.Cells(i, 7).Value
            Else
                vol_total = vol_total + ws.Cells(i, 7).Value
                close_price = ws.Cells(i, 6).Value
                ws.Cells(j, 9).Value = ticker
                delta_price = close_price - open_price
                If (open_price >= 0) Then
                    percent_change = (delta_price / open_price)
                Else
                    percent_change = 0
                End If
                ws.Cells(j, 10).Value = delta_price
                ws.Cells(j, 11).Value = percent_change
                ws.Cells(j, 12).Value = vol_total
                If (delta_price < 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red
                ElseIf (delta_price = 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 2 ' White
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 4 ' Green
                End If
                open_price = ws.Cells(i + 1, 3).Value
                ticker = ws.Cells(i + 1, 1).Value
                j = j + 1
                vol_total = 0
                delta_price = 0
            End If
        Next i
            
        Dim per_min As Double
        Dim per_max As Double
        Dim vol_max As Double
        Dim k As Double
        Dim row_count1 As Double
        Dim get_ticker As String
        per_min = Application.WorksheetFunction.Min(ws.Range("K:K"))
        per_max = Application.WorksheetFunction.Max(ws.Range("K:K"))
        vol_max = Application.WorksheetFunction.Max(ws.Range("L:L"))
        row_count1 = WorksheetFunction.CountIf(ws.Range("K:K"), "<>")
        For k = 2 To row_count1 - 1
            If (per_min = ws.Cells(k, 11).Value) Then
                get_ticker = ws.Cells(k, 9).Value
                ws.Cells(3, 15).Value = get_ticker
                ws.Cells(3, 16).Value = per_min
            End If
            If (per_max = ws.Cells(k, 11).Value) Then
                get_ticker = ws.Cells(k, 9).Value
                ws.Cells(2, 15).Value = get_ticker
                ws.Cells(2, 16).Value = per_max
            End If
            If (vol_max = ws.Cells(k, 12).Value) Then
                get_ticker = ws.Cells(k, 9).Value
                ws.Cells(4, 15).Value = get_ticker
                ws.Cells(4, 16).Value = vol_max
            End If
        Next k
    Next ws
End Sub


