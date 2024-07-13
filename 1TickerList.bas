Attribute VB_Name = "Module1"
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
        
        On Error GoTo 0
    Next sheetName
End Sub

