Sub homeAssignmentVB()
Dim ws As Worksheet

For Each ws In Worksheets

Dim totalVolume As Double
totalVolume = 0

Dim row As Double
row = 2

Dim openPrice As Double
openPrice = 2

Dim yearlyChange As Double
Dim percentChange As Double


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "% Change"
ws.Range("L1") = "Total Volume"

ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"

ws.Range("N2") = "% Max"
ws.Range("N3") = "% Min"
ws.Range("N4") = "Max Volume"





For i = 2 To lastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
'        openValue = ws.Cells(openPrice, 3).Value
        
        yearlyChange = (ws.Cells(i, 6).Value - ws.Cells(openPrice, 3).Value)
        
        
        
'  This is condition to check if division by zero.  Then assign Ticker open price "zero" value to avoid division by "zero"
        
        If ws.Cells(openPrice, 3).Value = 0 Then
            percentChange = 0
        Else
            percentChange = ((ws.Cells(i, 6).Value - ws.Cells(openPrice, 3).Value) / ws.Cells(openPrice, 3).Value)

        End If
        
' Putting the values in the Column Headings defined above; Ticker, Yearly Change, % Change and Total Volume

        ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(row, 10).Value = yearlyChange
        ws.Cells(row, 11).Value = percentChange
        ws.Cells(row, 11).NumberFormat = "0.00%"
        
        ws.Cells(row, 12).Value = totalVolume
        
'Conditional Formatting the Yearly Change
        
        If yearlyChange < 0 Then
            ws.Cells(row, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(row, 10).Interior.ColorIndex = 4
        End If

        
        totalVolume = 0
        row = row + 1
        openPrice = i + 1
       
    Else
    
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
    End If
    
Next i



' Greatest %age Increase
ws.Cells(2, 16).Value = WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(2, 16).NumberFormat = "0.00%"

'Greatest %age Decrease
ws.Cells(3, 16).Value = WorksheetFunction.Min(ws.Range("K:K"))
ws.Cells(3, 16).NumberFormat = "0.00%"

'Greatest Total Volume
ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))

' Return matching stockID to the Greatest %age Increase

ws.Cells(2, 15).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Cells(2, 16), ws.Range("K:K"), 0))

' Return matching stockID to the Greatest %age Decrease

ws.Cells(3, 15).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Cells(3, 16), ws.Range("K:K"), 0))

' Return matching stockID to the Greatest Total Volume

ws.Cells(4, 15).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Cells(4, 16), ws.Range("L:L"), 0))


Next ws

    

End Sub


