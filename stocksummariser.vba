Sub stocksummariser()
    
    ' set up headings from col 9 on
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' first get the unique list of tickers
    
    Dim nrows As Long
    nrows = Range("A1").End(xlDown).Row
    
    Dim pointer_row As Long
    pointer_row = 1
    
    For i = 2 To nrows
        If Not Cells(pointer_row, 9) = Cells(i, 1).Value Then
            pointer_row = pointer_row + 1
            Cells(pointer_row, 9) = Cells(i, 1).Value
        End If
        Application.StatusBar = "Generating ticker list: " & i & " out of " & nrows - 1 & " records parsed: " & Format(i / (nrows-1), "0%")
    Next i
    
    ' then iterate to obtain changes and volumes
     
    Dim ntickerows As Long
    ntickerrows = Cells(1, 9).End(xlDown).Row

    Dim xdate As Long
    Dim firstdate As Long
    Dim lastdate As Long
    Dim firstdaterow As Long
    Dim lastdaterow As Long
    Dim volume As LongLong
    Dim ticker As String
    Dim openval As Double
    Dim closeval As Double
    Dim change As Double
    Dim changepct As Double
    
    For i = 2 To ntickerrows
        
        ' initialise values per ticker
        ticker = Cells(i, 9).Value
        volume = 0
        firstdate = 99999999
        lastdate = 0
        
        ' run through all records to get volume and first/last days
        For j = 2 To nrows
            If Cells(j, 1).Value = ticker Then
                volume = volume + Cells(j, 7).Value
                xdate = Cells(j, 2).Value
                If xdate < firstdate Then
                    firstdate = xdate
                    firstdaterow = j
                End If
                If xdate > lastdate Then
                    lastdate = xdate
                    lastdaterow = j
                End If
            End If
        Next j
        
        ' calculate diffs
        openval = Cells(firstdaterow, 3).Value
        closeval = Cells(lastdaterow, 6).Value
        
        change = closeval - openval
        changepct = change / openval
        
        ' write the values to the table
        Cells(i, 10).Value = change
        Cells(i, 11).Value = changepct
        Cells(i, 12).Value = volume
        
        ' format
        Cells(i, 11).NumberFormat = "0.00%"
        If changepct > 0 Then Cells(i, 11).Interior.ColorIndex = 4
        If changepct < 0 Then Cells(i, 11).Interior.ColorIndex = 3
        
        ' status bar update
        Application.StatusBar = "Progress: " & i & " tickers out of " & ntickerrows - 1 & ": " & Format(i / (ntickerrows-1), "0%")
        
    Next i
    
    ' bonus summary table
        
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    Cells(2, 17).Value = "=max(K:K)"
    Cells(3, 17).Value = "=min(K:K)"
    Cells(4, 17).Value = "=max(L:L)"
    
    Cells(2, 16).Value = "=xlookup(Q2,K:K,I:I)"
    Cells(3, 16).Value = "=xlookup(Q3,K:K,I:I)"
    Cells(4, 16).Value = "=xlookup(Q4,L:L,I:I)"
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"

    Application.StatusBar = False
    
End Sub

