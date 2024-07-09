Attribute VB_Name = "Module33"
Sub HighestVolumeQ2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2") ' Set the worksheet to Q2

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row ' Find the last row in column L

    Dim HighestNumber As Double
    Dim TickerWithHighestNumber As String
    Dim i As Long

    If lastRow >= 2 Then
        HighestNumber = ws.Cells(2, 12).Value ' Initialize HighestNumber with the value from the first row in column L
        TickerWithHighestNumber = ws.Cells(2, 9).Value ' Initialize TickerWithHighestNumber with the corresponding Ticker value from column I

        ' Loop through values in column L starting from the second row
        For i = 3 To lastRow
            If ws.Cells(i, 12).Value > HighestNumber Then
                HighestNumber = ws.Cells(i, 12).Value ' Update HighestNumber if current value is higher
                TickerWithHighestNumber = ws.Cells(i, 9).Value ' Update TickerWithHighestNumber with the corresponding Ticker
            End If
        Next i
    Else
        MsgBox "No data found in column L."
        Exit Sub
    End If

    ' Output the highest number value in column L to Q4
    ws.Cells(4, 16).Value = HighestNumber ' Column Q, Row 4

    ' Output the corresponding Ticker value from column I to P4
    ws.Cells(4, 15).Value = TickerWithHighestNumber ' Column P, Row 4
End Sub


