Attribute VB_Name = "Module37"
Sub HighestPercentage()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim HighestPercentage As Double
    Dim TickerWithHighestPercentage As String
    Dim i As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row ' Find the last row in column K
        
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
            MsgBox "No data found in column K for sheet " & sheetName
            Exit Sub
        End If

        ' Output the highest percentage value in column K to Q2
        ws.Cells(2, 17).Value = HighestPercentage ' Column Q, Row 2

        ' Output the corresponding Ticker value from column I to P2
        ws.Cells(2, 16).Value = TickerWithHighestPercentage ' Column P, Row 2
    Next sheetName
End Sub
