Attribute VB_Name = "Module15"
Sub QuarterlyChangeQ2()
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2") ' Set the worksheet to Q2

    Dim lastRowA As Long
    Dim lastRowI As Long
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row in column A
    lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row ' Find the last row in column I

    Dim Ticker As String
    Dim CurrentTicker As String
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim firstRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long

    For j = 2 To lastRowI ' Loop through values in column I, starting from I2
        Ticker = ws.Cells(j, 9).Value ' Get the value from column I (I2, I3, etc.)

        ' Initialize variables
        OpenValue = 0
        CloseValue = 0
        firstRow = 0
        lastRow = 0

        For i = 2 To lastRowA ' Loop through values in column A, starting from row 2
            CurrentTicker = ws.Cells(i, 1).Value ' Get the value from column A

            If CurrentTicker = Ticker Then
                If firstRow = 0 Then
                    firstRow = i ' Store the first row where Ticker is found
                    OpenValue = ws.Cells(i, 3).Value ' Store the OpenValue from column C
                End If
                lastRow = i ' Update the last row where Ticker is found
                CloseValue = ws.Cells(i, 6).Value ' Update the CloseValue from column F
            End If
        Next i

        If firstRow > 0 And lastRow > 0 Then
            ws.Cells(j, 10).Value = CloseValue - OpenValue ' Subtract CloseValue from OpenValue and place the result in column J
        End If
    Next j
End Sub

