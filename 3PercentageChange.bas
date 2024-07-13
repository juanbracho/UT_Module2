Attribute VB_Name = "Module4"
Sub PercentageChange()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRowA As Long
    Dim Ticker As String
    Dim CurrentTicker As String
    Dim OpenValue As Double
    Dim ChangeValue As Double
    Dim i As Long
    Dim j As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row in column A

        For i = 2 To 1501 ' Loop through values in column I, from I2 to I1501
            Ticker = ws.Cells(i, 9).Value ' Get the value from column I (I2, I3, etc.)

            ' Initialize OpenValue
            OpenValue = 0

            For j = 2 To lastRowA ' Loop through values in column A, starting from row 2
                CurrentTicker = ws.Cells(j, 1).Value ' Get the value from column A

                If CurrentTicker = Ticker Then
                    OpenValue = ws.Cells(j, 3).Value ' Store the OpenValue from column C
                    Exit For ' Exit the loop once the first match is found
                End If
            Next j

            If OpenValue <> 0 Then
                ChangeValue = ws.Cells(i, 10).Value ' Get the value from column J
                If OpenValue <> 0 Then
                    ws.Cells(i, 11).Value = ChangeValue / OpenValue ' Divide ChangeValue by OpenValue and place the result in column K
                End If
            End If
        Next i
    Next sheetName
End Sub
