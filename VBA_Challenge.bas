Attribute VB_Name = "Module2"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Cells(3, 4).Value = "ticker Starting Price"
    Cells(3, 5).Value = "ticker Ending Price"
    
    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim index As Byte

    '1b) Create three output arrays
    Dim tickerVolumes As Double
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
   
   For index = 0 To 11
       ticker = tickers(index)
       tickerVolumes = 0
       
    ''2b) Loop over all the rows in the spreadsheet.
      Sheets(yearValue).Activate
        
        For i = 2 To RowCount
        
           '3a) Increase volume for current ticker
           If Cells(i, 1).Value = ticker Then

               tickerVolumes = tickerVolumes + Cells(i, 8).Value

           End If
           '3b) Check if the current row is the first row with the selected tickerIndex - ticker Starting Price
           If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

               tickerStartingPrices = Cells(i, 6).Value

           End If

           '3c) check if the current row is the last row with the selected tickerIndex -ticker Ending Price
           If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

               tickerEndingPrices = Cells(i, 6).Value

           End If
       Next i
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + index, 1).Value = ticker
       Cells(4 + index, 2).Value = tickerVolumes
       Cells(4 + index, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
       Cells(4 + index, 4).Value = tickerStartingPrices
       Cells(4 + index, 5).Value = tickerEndingPrices
    
Next index
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
