Attribute VB_Name = "Module24"
Sub TotalStockQ3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q3") ' Set the worksheet to Q3

    Dim lastRowA As Long
    Dim lastRowI As Long
    Dim i As Long
    Dim j As Long
    Dim Ticker As String
    Dim CurrentTicker As String
    Dim SumValue As Double

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row in column A
    lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row ' Find the last row in column I

    ' Loop through the range in column I
    For i = 2 To lastRowI
        Ticker = ws.Cells(i, 9).Value ' Get the value from column I (I2, I3, etc.)
        SumValue = 0 ' Initialize the sum value

        ' Loop through the range in column A
        For j = 2 To lastRowA
            CurrentTicker = ws.Cells(j, 1).Value ' Get the value from column A

            ' If the value in column I matches the value in column A, sum the values in column G
            If CurrentTicker = Ticker Then
                SumValue = SumValue + ws.Cells(j, 7).Value
            End If
        Next j

        ' Return the sum result in column L
        ws.Cells(i, 12).Value = SumValue ' Column L corresponds to column 12
    Next i
End Sub

