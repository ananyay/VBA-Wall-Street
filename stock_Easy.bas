Attribute VB_Name = "stock_Easy"
Sub stock_easy()
Dim i As Long
Dim resItr As Integer
Dim lrow As Long

For Each ws In ActiveWorkbook.Worksheets

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Total Volume"
    
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    resItr = 2
    
    For i = 2 To lrow
        Volume = Volume + ws.Cells(i, 7).Value
       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ws.Cells(resItr, 9).Value = ticker
            ws.Cells(resItr, 10).Value = Volume
            Volume = 0
            resItr = resItr + 1
        End If
    Next i
Next ws



End Sub
