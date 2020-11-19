Attribute VB_Name = "Module3"
Sub Formatting():
For Each ws In Worksheets
    
    Dim lastrow As Long
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' Establishes variable names to be placed in every sheet for
    ' the new variables added.
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    
    ' for loop to run through the new data that has been compiled
    ' to search for positive values and color the cell green
    ' and look for negative values and color the cell red.
    For i = 2 To lastrow
    
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 0
    
        End If
    
    Next i
 
    
Next ws


End Sub
