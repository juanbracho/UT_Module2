Attribute VB_Name = "Module5"
Sub CombinedScript()
    Call TickerList
    Call QuarterlyChange
    Call PercentageChange
    Call TotalStock
    Call HighestPercentage
    Call LowestPercentage
    Call HighestVolume
End Sub
Sub TickerList()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim uniqueValues As Collection
    Dim currentValue As String
    Dim i As Long
    Dim outputRow As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        Set uniqueValues = New Collection
        
        outputRow = 2 ' Start output in I2
        
        On Error Resume Next
        
        For i = 2 To lastRow
            currentValue = ws.Cells(i, 1).Value
            
            If Len(currentValue) > 0 Then
                uniqueValues.Add currentValue, CStr(currentValue)
                
                If Err.Number = 0 Then
                    ws.Cells(outputRow, 9).Value = currentValue
                    outputRow = outputRow + 1
                Else
                    Err.Clear
                End If
            End If
        Next i
    Next sheetName
End Sub

Sub QuarterlyChange()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRowA As Long
    Dim lastRowI As Long
    Dim Ticker As String
    Dim CurrentTicker As String
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim firstRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row in column A
        lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row ' Find the last row in column I

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
    Next sheetName
End Sub

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

Sub TotalStock()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRowA As Long
    Dim lastRowI As Long
    Dim i As Long
    Dim j As Long
    Dim Ticker As String
    Dim CurrentTicker As String
    Dim SumValue As Double
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
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
    Next sheetName
End Sub

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
        ws.Cells(3, 17).Value = LowestPercentage ' Column Q, Row 3

        ' Output the corresponding Ticker value from column I to P3
        ws.Cells(3, 16).Value = TickerWithLowestPercentage ' Column P, Row 3
    Next sheetName
End Sub

Sub HighestVolume()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim HighestNumber As Double
    Dim TickerWithHighestNumber As String
    Dim i As Long
    
    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row ' Find the last row in column L
        
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
            MsgBox "No data found in column L for sheet " & sheetName
            Exit Sub
        End If

        ' Output the highest number value in column L to Q4
        ws.Cells(4, 17).Value = HighestNumber ' Column Q, Row 4

        ' Output the corresponding Ticker value from column I to P4
        ws.Cells(4, 16).Value = TickerWithHighestNumber ' Column P, Row 4
    Next sheetName
End Sub

