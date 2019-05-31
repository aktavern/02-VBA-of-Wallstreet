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
    Dim greatestIncrease As Double
    Dim greatestSymbol As String
    Dim greatestDecrease As Double
    Dim decreaseSymbol As String
    Dim greatestVolume As Double
    Dim volumeSymbol As String
    
    'loop through each row in the spreadsheet
    For Each ws In Worksheets
     ' set titles for columns in each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ' set variables at begining of each worksheet
        greatestIncrease = 0
        greatestSymbol = ""
        greatestDecrease = 0
        decreaseSymbol = ""
        greatestVolume = 0
        volumeSymbol = ""
        tickerSymbol = ""
        totalVolume = 0
        rowNumber = 1
        cnt = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            cnt = cnt + 1
            ' check if the next row is the same as the current row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                'update variables with current data with current data if rows don't match
                tickerSymbol = ws.Cells(i, 1).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                rowNumber = rowNumber + 1
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
                ' determine greatest increase or decrease
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestSymbol = tickerSymbol
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseSymbol = tickerSymbol
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeSymbol = tickerSymbol
                End If
                'update values of summary tables
                ws.Range("I" & rowNumber).Value = tickerSymbol
                ws.Range("J" & rowNumber).Value = yearlyChange
                ws.Range("K" & rowNumber).Value = Format(percentChange, "Percent")
                ws.Range("L" & rowNumber).Value = totalVolume
                ws.Cells(2, 16).Value = greatestSymbol
                ws.Cells(2, 17).Value = Format(greatestIncrease, "Percent")
                ws.Cells(3, 16).Value = decreaseSymbol
                ws.Cells(3, 17).Value = Format(greatestDecrease, "Percent")
                ws.Cells(4, 16).Value = volumeSymbol
                ws.Cells(4, 17).Value = greatestVolume
                'reset variables
                totalVolume = 0
                cnt = 0
                openValue = 0
                closeValue = 0
                percentChange = 0
                tickerSymbol = ""
            'add to total volume if row values are the same
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        'reset variables for next iteration
        rowNumber = 1
        cnt = 0
        totalVolume = 0
    Next ws
    
End Sub


















