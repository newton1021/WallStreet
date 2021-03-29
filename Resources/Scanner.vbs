sub scanner()
    'move to the start location in col A and row 2
    dim curRow as Integer
    dim openPrice as double
    dim closePrice as double
    dim volumn as Integer
    dim symbol as string
    dim change as double

    dim fStop as Boolean

    curRow = 2
    fStop = false
    while not fStop 
        'get the ticker symbol
        if symbol = "" then
            symbol = cells(curRow, 1).value
            openPrice = cells(curRow, 3).value
            volumn = 0
        end if

        volumn = volumn + cells(curRow, 7).value
        'get date of transaction
        'is this the first date? or last date
        'get high
        'get low
        'get close
        'get volumn Summing

        if cells(curRow + 1, 1).value = "" Then
            fStop = true
        end if

        if cells(curRow + 1, 1) <> symbol then 
            'finished the symbol report results.
            closePrice = cells(curRow, 6).value
            change = closePrice - openPrice
            percent = change / openPrice

            msgBox(symbol & " change " & change & " (" & percent * 100 & "% vol:" & volumn)

            symbol = ""
        end if

        curRow = curRow + 1
    wend


end Sub