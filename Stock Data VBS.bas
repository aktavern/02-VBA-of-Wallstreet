Attribute VB_Name = "Module1"
Sub stockdata()
    ' define variables to be used in loop
    Dim tickerSymbol As String
    Dim totalVolume As Double
    Dim rowNumber As Double
    Dim lastRow As Long
    Dim cnt As Double
    Dim openValue As Double
    Dim closeValue As Double
    Dim percentChange As Double
    Dim yearlyChange As Double
    
    ' set starter values for variables
    tickerSymbol = ""
    totalVolume = 0
    rowNumber = 1
    cnt = 0
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'loop through each row in the spreadsheet
    For Each ws In Worksheets
     ' set titles for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        For i = 2 To lastRow
            cnt = cnt + 1
            ' check if the next row is the same as the current row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                'update variables with current data with current data if rows don't match
                tickerSymbol = ws.Cells(i, 1).Value
                rowNumber = rowNumber + 1
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                openValue = ws.Cells(Abs(i - cnt) + 1, 3)
                closeValue = ws.Cells(i, 6)
                yearlyChange = closeValue - openValue
                'format yearly change
                If yearlyChange > 0 Then
                    ws.Range("J" & rowNumber).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & rowNumber).Interior.ColorIndex = 3
                End If
                ' prevent divide by 0 errors when evaluating percent change
                If openValue = 0 Then
                    ws.Range("L" & rowNumber).Value = 0
                Else
                    percentChange = ((closeValue / openValue) - 1)
                End If
                'update values of summary table
                ws.Range("I" & rowNumber).Value = tickerSymbol
                ws.Range("J" & rowNumber).Value = yearlyChange
                ws.Range("K" & rowNumber).Value = Format(percentChange, "Percent")
                ws.Range("L" & rowNumber).Value = totalVolume
                'reset variables
                totalVolume = 0
                cnt = 0
                openValue = 0
                closeValue = 0
                percentChange = 0
            'add to total volume if row values are the same
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
                tickerSymbol = ""
        Next i
        rowNumber = 1
        cnt = 0
        totalVolume = 0
    Next ws
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
End Sub













