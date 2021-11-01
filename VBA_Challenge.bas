Attribute VB_Name = "Module1"
Sub macrocheck()

    Dim testMessage As String
    
    testMessage = "Hello World!"
    
    MsgBox (testMessage)
    
End Sub


Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    'establish number fo rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    'set initial volume to 0
    totalVolume = 0
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists

    Dim startingPrice As Double
    Dim endingPrice As Double
    
    
    'Loop over all the rows
    For i = rowStart To rowEnd
    
        'increase totalVolume if ticker is "DQ" by the value in the current row
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
    
        'Find first instance of DQ to set starting price
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If
        
        'Find last instance of DQ to set ending price
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If
                
    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    

End Sub

Sub AllStocksAnalysis()

'1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    'Title the worksheet
    Range("A1").Value = "All Stocks(2018)"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
 

'2) Initialize an array of all tickers.
    Dim tickers(11) As String

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

'3) Prepare for the analysis of tickers.

'3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double

'3b) Activate the data worksheet.
    Worksheets("2018").Activate

'3c) Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
    '5) Loop through rows in the data.
        Worksheets("2018").Activate
        For j = 2 To RowCount
        
        '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
        '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
        '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
    '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


    Next i
    
    

    
    
End Sub


Sub formatAllStocksAnalysisTable()

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("a3:c3").Font.Bold = True
    Range("a3:c3").Font.Size = 14
    Range("a3:c3").Font.Color = RGB(0, 255, 255)
    
    Range("a3:c3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("b4:b15").NumberFormat = "$#,##0.00"
    Range("c4:c15").NumberFormat = "0.00%"
    
    Columns("B").AutoFit
    
    'Conditional formatting for Return column
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
            
        'Clear the cell color
        Else: Cells(i, 3).Interior.Color = xlNone
        
        End If
        
    Next i
    

End Sub


Sub ClearWorksheet()

    Worksheets("All Stocks Analysis").Activate

    Cells.Clear

End Sub



Sub ClearDQWorksheet()

    Worksheets("DQ Analysis").Activate

    Cells.Clear

End Sub


Sub yearValueAnalysis()

    'Find speed of VBA code
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

'1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    'Title the worksheet
    Range("A1").Value = "All Stocks(" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
 

'2) Initialize an array of all tickers.
    Dim tickers(11) As String

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

'3) Prepare for the analysis of tickers.

'3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double

'3b) Activate the data worksheet.
    Worksheets(yearValue).Activate

'3c) Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
    '5) Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
        '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
        '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
    '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Create a sheet title
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerindex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j

    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
                            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        End If
    
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerindex = tickerindex + 1
        End If

    Next i
   
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        'Format positive and negative returns
        If Cells(4 + i, 3).Value > 0 Then
            Cells(4 + i, 3).Interior.Color = vbGreen
        
        Else
            Cells(4 + i, 3).Interior.Color = vbRed
        
        End If
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub




































