Sub AllStocksAnalysisRefactored()

'    ' turn off these features
'    SwitchOff (True)
    
    'Our new macro should do the following:
    
    'Timing
    Dim startTime As Single, endTime As Single
    
    'InputBox value to a variable then to a string.
    Dim yearValue As Variant
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Start Timer
    startTime = Timer
    
    '1. Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2. Initialize an array of all tickers and ticket temp var.
    Dim tickers(12) As String, ticker As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

'3. Prepare for the analysis of tickers.
    
    'Initialize variables for the starting price and ending price.
    Dim totalVolume As Single, startingPrice As Single, endingPrice As Single
    
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'find the number of rows to loop over
    Dim RowCount As Integer
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Set ticker index to 0
    Dim tickerIndex As Integer, tickerCount As Integer
    tickerIndex = 0: tickerCount = 11

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Type out the variables we are using for loops
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    
    '2a) Initialize ticker volumes to zero
    For i = 0 To tickerCount
        
        tickerVolumes(i) = 0
        
    Next i

    '2b) Activiate Worksheet and loop over all the rows
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

            
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then

                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
            End If
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            '3d Increase the tickerIndex if the next rows ticker does not match the last rows ticker.
                tickerIndex = tickerIndex + 1
                
            End If
        
    Next j

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For k = 0 To tickerCount
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumes(k)
        Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
        
    Next k
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    'https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-with-statements
    With Range("A3:C3")
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    Dim dataRowStart As Integer, dataRowEnd As Integer
    dataRowStart = 4: dataRowEnd = 15

    For l = dataRowStart To dataRowEnd
        
        If Cells(l, 3) > 0 Then
            
            Cells(l, 3).Interior.Color = vbGreen
            
        Else

            Cells(l, 3).Interior.Color = vbRed
            
        End If
        
    Next l
    
    'End Timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

'    ' turn these features back on
'    SwitchOff (False)

End Sub