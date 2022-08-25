# VBA_Challenge.vbs-
Module 2 Homework
 # Summry report 
 The purpose of this analysis is so that the client Steve can see an expanded version of the dataset to include the entire stock market over the last few years. There are thousands of stocks that need to be looked at. Where going to be giving him version of the data that compares the year of 2017 and 2018. 

I’ll be executing the original script and the refactored script.  We will be trying to refractor the code and loop in all the spreadsheet, find a ticker, we use then use to retrieve the total volume of the stock by using its starting and ending of the stocks themselves.  We will also make sure that the code successfully made the VBA script run faster.
# Here is the Code 
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
    'Use long and single
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerStartingPrices(i) = 0
        Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
       
       
       
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        
        
        
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            
      
            
            '3d Increase the tickerIndex.
            'Add plus 1
            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
        End If

        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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
 
End Sub

<img width="259" alt="Screenshot 2022-08-25 105221" src="https://user-images.githubusercontent.com/109318020/186698339-9517820d-aad9-48bd-8938-93e9fd463b8e.png">



<img width="345" alt="Screenshot 2022-08-25 103536" src="https://user-images.githubusercontent.com/109318020/186696944-102f9dce-e7dd-4b36-9f61-b7f4a0fc265b.png">


<img width="258" alt="VBA_Challenge_2017 png " src="https://user-images.githubusercontent.com/109318020/186702959-1d99b326-8927-4e5e-8021-f8cafb86c10e.png">

<img width="251" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/109318020/186703031-993532ff-df26-4bf3-9eb8-d256eeab5fa1.png">

