Attribute VB_Name = "Module12"
Sub TickerListQ2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim uniqueValues As Collection
    Set uniqueValues = New Collection
    
    Dim currentValue As String
    Dim i As Long
    Dim outputRow As Long
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
    
    On Error GoTo 0
End Sub
