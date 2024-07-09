Attribute VB_Name = "Module28"
Sub HighestPercentageQ4()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q4") ' Set the worksheet to Q4

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row ' Find the last row in column K

    Dim HighestPercentage As Double
    Dim TickerWithHighestPercentage As String
    Dim i As Long

    If lastRow >= 2 Then
        HighestPercentage = ws.Cells(2, 11).Value ' Initialize HighestPercentage with the first value in column K
        TickerWithHighestPercentage = ws.Cells(2, 9).Value ' Initialize TickerWithHighestPercentage with the corresponding value in column I

        ' Loop through values in column K starting from the second row
        For i = 3 To lastRow
            If ws.Cells(i, 11).Value > HighestPercentage Then
                HighestPercentage = ws.Cells(i, 11).Value ' Update HighestPercentage if current value is higher
                TickerWithHighestPercentage = ws.Cells(i, 9).Value ' Update TickerWithHighestPercentage with the corresponding Ticker
            End If
        Next i
    Else
        MsgBox "No data found in column K."
        Exit Sub
    End If

    ' Output the highest percentage value in column K to Q2
    ws.Cells(2, 16).Value = HighestPercentage ' Column Q, Row 2

    ' Output the corresponding Ticker value from column I to P2
    ws.Cells(2, 15).Value = TickerWithHighestPercentage ' Column P, Row 2
End Sub

