Attribute VB_Name = "Module6"
Sub LowestPercentage()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim LowestPercentage As Double
    Dim TickerWithLowestPercentage As String
    Dim i As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row ' Find the last row in column K
        
        If lastRow >= 2 Then
            LowestPercentage = ws.Cells(2, 11).Value ' Initialize LowestPercentage with the value from the first row in column K
            TickerWithLowestPercentage = ws.Cells(2, 9).Value ' Initialize TickerWithLowestPercentage with the corresponding Ticker value from column I

            ' Loop through values in column K starting from the second row
            For i = 3 To lastRow
                If ws.Cells(i, 11).Value < LowestPercentage Then
                    LowestPercentage = ws.Cells(i, 11).Value ' Update LowestPercentage if current value is lower
                    TickerWithLowestPercentage = ws.Cells(i, 9).Value ' Update TickerWithLowestPercentage with the corresponding Ticker
                End If
            Next i
        Else
            MsgBox "No data found in column K for sheet " & sheetName
            Exit Sub
        End If

        ' Output the lowest percentage value in column K to Q3
        ws.Cells(3, 16).Value = LowestPercentage ' Column Q, Row 3

        ' Output the corresponding Ticker value from column I to P3
        ws.Cells(3, 15).Value = TickerWithLowestPercentage ' Column P, Row 3
    Next sheetName
End Sub

