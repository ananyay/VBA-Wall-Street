Attribute VB_Name = "Stock_hard"
Sub stock_hard():
' Module for the easy portion of the assignment

' Variables
' Variable to store the row count in the sheet
Dim lastRow As Long
' Variable to record the current ticker we are processing
Dim ticker As String
' Varible to record the total volume for each ticker
Dim Volume As Variant
' Variable to capture the Yearly opening cost
Dim yearOpen As Double
' Variable to capture the yearly closing cost
Dim yearClose As Double
' Variable to capture the yearly change
Dim yearChange As Single
' Variable to capture the yearly percent change
Dim pctChange As Variant
'Counter Variable to parse the rows in the Main table
Dim rowCounter As Variant
'Counter Variable to parse the rows in the Summary Table
Dim resCounter As Variant



' parse each sheet in the workbook
For Each ws In ActiveWorkbook.Worksheets
    ' Assign header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock volume"
        
    ' initialize the counter for summary table
    resCounter = 2
    
    ' Determine last row and update to the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Initialize volume variable as 0
    Volume = 0
    'Initialize openining cost for the first ticker in this worksheet
    yearOpen = ws.Cells(2, 3).Value
    
    ' Parse for each row
    For rowCounter = 2 To lastRow
        Volume = Volume + ws.Cells(rowCounter, 7).Value
                
        'If this condition succeeds, it means we are at the end of the ticker.
           If (ws.Cells(rowCounter - 1, 1).Value = ws.Cells(rowCounter, 1).Value And ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
            yearClose = ws.Cells(rowCounter, 6).Value
            yearChange = yearClose - yearOpen
           ' If the opening cost is 0
           If yearOpen = 0 Then
            pctChange = (yearClose - yearOpen) / 100
           Else
            pctChange = (yearClose - yearOpen) / yearOpen
           End If
           yearOpen = ws.Cells(rowCounter + 1, 3).Value
           ' write the year change to the worksheet
           ws.Cells(resCounter, 10).Value = yearChange
           ' Conditionally format the yearly change
           If ws.Cells(resCounter, 10).Value < 0 Then
            ws.Cells(resCounter, 10).Interior.ColorIndex = 3
           Else
            ws.Cells(resCounter, 10).Interior.ColorIndex = 4
           End If
           
           ' write the percent change to the worksheet
           ws.Cells(resCounter, 11).Value = pctChange
           ' Reformat percent change as percentage
           ws.Cells(resCounter, 11).NumberFormat = "0.00%"
           
        End If
        
        ' if this condition succeeds, it means that the next ticker has changed.
        ' Record the total volume for this ticker in the summary table
        ' Record the ticker name for this ticker in the summary table
        If (ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
            ' set the ticker value in the summary table
            ws.Cells(resCounter, 9).Value = ws.Cells(rowCounter, 1).Value
            ws.Cells(resCounter, 12).Value = Volume
        
            'Reset the volume for next ticker
            Volume = 0
        
            ' Increment the counter for summary table
           resCounter = resCounter + 1
        End If
        
    Next rowCounter

    ' This section is for the hard part
    ' Reset the variables to start parsing the summary table again
    resCounter = 2
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Assign headers again
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Initialize these cells to 0, so that we can start comparing
    
    ws.Cells(2, 17).Value = 0
    ws.Cells(3, 17).Value = 0
    ws.Cells(4, 17).Value = 0
    For resCounter = 2 To lastRow
        ' For % increase, it is 11
        If ws.Cells(resCounter, 11).Value > ws.Cells(2, 17).Value Then
            ws.Cells(2, 17).Value = ws.Cells(resCounter, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(resCounter, 9).Value
        End If
        ' For % increase, it is 11
        If ws.Cells(resCounter, 11).Value < ws.Cells(3, 17).Value Then
            ws.Cells(3, 17).Value = ws.Cells(resCounter, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(resCounter, 9).Value
        End If
        ' For total volume, it is 12
         If ws.Cells(resCounter, 12).Value > ws.Cells(4, 17).Value Then
            ws.Cells(4, 17).Value = ws.Cells(resCounter, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(resCounter, 9).Value
        End If
        
    Next resCounter
    ' Reformat percent change as percentage
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "General"
Next ws

End Sub




