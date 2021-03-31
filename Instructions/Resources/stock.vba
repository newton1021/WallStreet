Dim ResultRow As Long
Dim results As Worksheet

Sub scanner()
    ' This function will scan through the current worksheet and exract the data to the results page.
    Dim curRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim symbol As String
    Dim change As Double
    Dim yDate As String
    Dim fStop As Boolean
    
    'initialize for the page.
    openPrice = 0
    curRow = 2
    fStop = False
    ResultRow = 2
    'continue until reaching the end of the rows.
    While Not fStop
        'get the ticker symbol if the symbole is "" then that is the end of the data for this sheet.
        If symbol = "" Then
            symbol = Cells(curRow, 1).Value
            openPrice = Cells(curRow, 3).Value
            volume = 0
        End If

        volume = volume + Cells(curRow, 7).Value
        'get date of transaction
        'is this the first date? or last date
        'get high
        'get low
        'get close
        'get volume Summing

        
        
        
        
        If Cells(curRow + 1, 1).Value = "" Then
            fStop = True
        End If

        If Cells(curRow + 1, 1) <> symbol Then
            yDate = Left(Cells(curRow, 2).Value, 4)
            'finished the symbol report results.
            closePrice = Cells(curRow, 6).Value
            change = closePrice - openPrice
            If openPrice <> 0 Then
            Percent = change / openPrice
            Else
            Percent = 0
            End If
            'MsgBox (symbol & " change " & change & " (" & Percent * 100 & "% vol:" & volume)
            startCol = 10
            results.Cells(ResultRow, startCol + 1) = symbol
            results.Cells(ResultRow, startCol + 2) = yDate
            results.Cells(ResultRow, startCol + 3) = openPrice
            results.Cells(ResultRow, startCol + 4) = closePrice
            results.Cells(ResultRow, startCol + 5) = change
            results.Cells(ResultRow, startCol + 6) = Percent
            results.Cells(ResultRow, startCol + 7) = volume
            ResultRow = ResultRow + 1

            symbol = ""
        End If

        curRow = curRow + 1
    Wend
    'MsgBox ("End: " & curRow)

End Sub

Sub StockYearReview()
    Dim sheet As Worksheet
   
    Dim fFound As Boolean
    
    fFound = False
    For Each sheet In Worksheets
        If sheet.Name = "Result" Then
            fFound = True
        End If
    Next sheet
    
    
    
    If Not fFound Then
        Sheets.Add.Name = "Result"
    End If
    
    Set results = Sheets("Result")
    results.Cells.Clear
    results.Range("A1:G1") = VBA.Array("Ticker", "Year", "Open", "Close", "Change", "Percent Change", "Volume")
    results.Range("A1:G1").Font.Bold = True
    results.Range("A1:G1").Interior.ColorIndex = 1
    results.Range("A1:G1").Font.ColorIndex = 2
    results.Range("F:F").NumberFormat = "% 0.0"
    With results.Range("E2:E" & Rows.Count)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(230, 50, 50)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(50, 230, 50)

    End With




    Call findSummary
End Sub


Sub findSummary()
    Sheets("Result").Select
    'Find the greatest gain/loss and maxVolumn

    Dim rowCount As Long
    Dim max As Double
    Dim min As Double
    Dim maxVolume As Double
    Dim maxTicker As String
    Dim mimTicker As String
    Dim volumeTicker As String

    rowCount = Cells(Rows.Count, 1).End(xlUp).Row



    max = 0
    min = 0
    maxVolume = 0

    maxTicker = ""
    minTicker = ""
    volumeTicker = ""


    For i = 2 To rowCount
        If Cells(i, 6) > max Then
            max = Cells(i, 6)
            maxTicker = Cells(i, 1)
        End If
        If Cells(i, 6) < min Then
            min = Cells(i, 6)
            minTicker = Cells(i, 1)
        End If
        If Cells(i, 7) > maxVolume Then
            maxVolume = Cells(i, 7)
            volumeTicker = Cells(i, 1)
        End If
    Next i
    Range("J2").Value = "Greatest % Increase"
    Range("J3").Value = "Greatest % Decrease"
    Range("J4").Value = "Greatest Total Volume"

    Range("K1").Value = "Ticker"
    Range("k2").Value = maxTicker
    Range("k3").Value = minTicker
    Range("K4").Value = volumeTicker
    
    Range("L1").Value = "Value"
    Range("L2").Value = max
    Range("L2").NumberFormat = "% #0.0"
    Range("L3").Value = min
    Range("L3").NumberFormat = "% #0.0"
    Range("L4").Value = maxVolume

End Sub


