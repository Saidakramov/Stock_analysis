# Stock_analysis

# _Overview of project_

  The purpose of this analysis was to determine which stocks returned maximum profits. We had 11 stocks and two years 2017 and 2018. Also we needed to refractor the codes in order for the VBA script to run smoother.

# _Results_

  ! [Image for 2017 results!](https://github.com/Saidakramov/Stock_analysis/blob/91c60929d0d3f544c3744f83da344ab8ba5e60d7/Screen%20Shot%202022-07-02%20at%205.28.13%20PM.png) 
  
  ! [Image for 2018 results!](https://github.com/Saidakramov/Stock_analysis/blob/91c60929d0d3f544c3744f83da344ab8ba5e60d7/Screen%20Shot%202022-07-02%20at%205.28.42%20PM.png)
  
  
  ## Original code
  

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single


   '1) Format the output sheet on All Stocks Analysis worksheet
        yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
        
    
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
    Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
        Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub


## Refractored code
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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
            
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then

               tickerStartingPrice(tickerIndex) = Cells(j, 6).Value
            
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then

                tickerEndingPrice(tickerIndex) = Cells(j, 6).Value

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
        'End If
            End If
    Next j
    

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
    Next i
    
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
 

  
  
  According to the results of 2017, we can see that most stocks had good total daily volumes and positive returns on all stocks except one, which is TERP, with -7.2%. DQ and SEDG have brought the most returns with 199.4% and 184.5%, respectively. 
  Refactored code ran in 0.0703125 seconds for the year 2017 at the time of running the code, while the original code ran in 0.2734375 seconds. Our refactored code proves to be running faster than the original code, which saves time and memory for the computer. 
  
  For the year 2018, we can witness that compared to the previous year, stocks performed less well. 2018 had a good total daily volume; however, only two stocks were able to return on investments with 81.9% and 84% for stocks ENPH and RUN respectively. All the other stocks provided negative returns.
  Refactored code ran in 0.07421875 seconds for the year 2018 at the time of running the code, and the original code ran in 0.2734375 seconds, which also proves to be slower than the modified code. 
  
  
  # _Summary_
  
  According to our refactored code, running time proves to be faster than the original code. This is one of the advantages of refactoring the codes, as well as code becoming more clean and neat. On the other hand, refactoring could become too complicated and confusing. 
  
  Orignal code, despite being slower, is still providing the results; that is one of the advantages. However, it will take a longer time for a computer to compute, and the code will look messier. 
  
  
  
  
