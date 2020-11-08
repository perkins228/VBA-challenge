Attribute VB_Name = "Module1"
Sub stocks()
Dim ws As Worksheet

For Each ws In Worksheets
    Dim i As Long
    Dim ticker As String
    Dim total As Double
    total = 0
    Dim chartrow As Integer
    chartrow = 2
    Dim lastrow As Long
    Dim lastticker As Long

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "% Changes"
    ws.Range("M1") = "Total Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"


    For i = 2 To lastrow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ticker = ws.Cells(i, 1)
            close_price = ws.Cells(i, 6)
            total = total + ws.Cells(i, 7)
            ws.Range("J" & chartrow) = ticker
            ws.Range("K" & chartrow) = close_price - open_price
                If (close_price - open_price) > 0 Then
                     ws.Range("K" & chartrow).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & chartrow).Interior.ColorIndex = 3
                End If
            ws.Range("L" & chartrow).NumberFormat = "0.00%"
                If open_price = 0 Then
                    ws.Range("L" & chartrow) = 0
                Else
                    ws.Range("L" & chartrow) = ((close_price / open_price) - 1)
                End If
            ws.Range("M" & chartrow) = total
            total = 0
            chartrow = chartrow + 1
        ElseIf ws.Cells(i, 2) = WorksheetFunction.Min(ws.Range("B2:B" & i)) Then
            open_price = ws.Cells(i, 3)
        Else
            total = total + ws.Cells(i, 7)

        End If

    Next i
    
    lastticker = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For i = 2 To lastticker
        If ws.Cells(i, 12) = WorksheetFunction.Max(ws.Range("L2:L" & lastticker)) Then
            ticker = ws.Cells(i, 10)
            greatest = ws.Cells(i, 12)
            ws.Range("P2") = ticker
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q2") = greatest
        End If
        
        If ws.Cells(i, 12) = WorksheetFunction.Min(ws.Range("L2:L" & lastticker)) Then
            ticker = ws.Cells(i, 10)
            lowest = ws.Cells(i, 12)
            ws.Range("P3") = ticker
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q3") = lowest
        End If
        
        If ws.Cells(i, 13) = WorksheetFunction.Max(ws.Range("M2:M" & lastticker)) Then
            ticker = ws.Cells(i, 10)
            largest = ws.Cells(i, 13)
            ws.Range("P4") = ticker
            ws.Range("Q4") = largest
        End If
    Next i

Next ws

End Sub

