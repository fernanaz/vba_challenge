Sub Stock_data():
Dim total_vol As Double
Dim last_row As Long
Dim Ticker_Symbol As String
Dim startingprice As Double
Dim endingprice As Double
Dim Summary_Table_Row As Integer
Dim percent_change As Double
Dim greatest_volume As Double
Dim vol_ticker As String
Dim greatest_percent As Double
Dim gp_ticker As String
Dim decrease_percent As Double
Dim dp_ticker As String
For Each ws In Worksheets
Summary_Table_Row = 2
total_vol = 0
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For I = 2 To last_row
        'if the next row is a new symbol
        If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
            startingprice = ws.Cells(I, 3).Value
        End If
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            'giving ticker symbol a value
            Ticker_Symbol = ws.Cells(I, 1).Value
            'total_vol value
            total_vol = total_vol + ws.Cells(I, 7).Value
            'ending price
            endingprice = ws.Cells(I, 6).Value
            Else
            total_vol = total_vol + ws.Cells(I, 7).Value
            End If
            
            If endingprice <> 0 Then
            'percent change
            percent_change = ((endingprice - startingprice) / endingprice)
            'printing in table
            ws.Range("I" & Summary_Table_Row) = Ticker_Symbol
            ws.Range("L" & Summary_Table_Row) = total_vol
            ws.Range("J" & Summary_Table_Row) = endingprice - startingprice
            ws.Range("K" & Summary_Table_Row) = percent_change
            'reset summary_table_row and total_vol
            Summary_Table_Row = Summary_Table_Row + 1
            total_vol = 0
            startingprice = 0
            endingprice = 0
        Else
        percent_change = 0
        End If
    Next I
'Adding Headings to Summary Table
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume)"
'Formatting colors for changes
For I = 2 To last_row
    If ws.Cells(I, 10) > 0 Then
    ws.Cells(I, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(I, 10) < 0 Then
    ws.Cells(I, 10).Interior.ColorIndex = 3
    Else
    End If
Next I
'formatting as a percent
For I = 2 To last_row
    ws.Cells(I, 11) = Format(ws.Cells(I, 11), "Percent")
Next I
'Challenge
'Making Headings for Table
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"
'Greatest % Increase
greatest_percent = ws.Cells(2, 11).Value
For I = 2 To last_row
    If ws.Cells(I, 11).Value > greatest_percent Then
   greatest_percent = ws.Cells(I, 11).Value
    gp_ticker = ws.Cells(I, 9).Value
    End If
Next I
ws.Range("P2") = greatest_percent
ws.Range("O2") = gp_ticker
ws.Range("P2") = Format(Range("P2"), "Percent")
'Greatest % Decrease
decrease_percent = ws.Cells(2, 11).Value
For I = 2 To last_row
    If ws.Cells(I, 11).Value < decrease_percent Then
   decrease_percent = ws.Cells(I, 11).Value
    dp_ticker = ws.Cells(I, 9).Value
    End If
Next I
ws.Range("P3") = decrease_percent
ws.Range("O3") = dp_ticker
ws.Range("P3") = Format(ws.Range("P3"), "Percent")
'Greatest Volume
greatest_volume = ws.Cells(2, 12).Value
For I = 2 To last_row
    If ws.Cells(I, 12).Value > greatest_volume Then
   greatest_volume = ws.Cells(I, 12).Value
   vol_ticker = ws.Cells(I, 9).Value
    End If
Next I
ws.Range("P4") = greatest_volume
ws.Range("O4") = vol_ticker
Next ws
End Sub