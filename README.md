# Analyzing Stocks in VBA
## Overview of Project
### Purpose
The main purpose for this project was to refactor the code generated through the module that collected information for a 12 stocks, but to do so in a faster time. This will be helpful since Steve's goal is to be able to run the code for a dataset that will have thousands of stocks, and not just 12.
## Results
The results for the original code for 2017 had an elapsed time of .5800781 seconds and for 2018 was .5878906 seconds.

By refactoring the below code, that was provided for this project in a VBS file, it help improved the time. Below the code you can also see the new times after the new code was written and run. 

    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
                      
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
             
            End If
         
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
For 2017, the above mentioned refactored code genearated a time of .1152344 seconds.

![VBA_Challenge_2017](https://github.com/DanielGandia/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

For 2018, the refactored code generated a time of .1210938 seconds. 

![VBA_Challenge_2018](https://github.com/DanielGandia/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
### Advantage of the Refactored Code
As previously stated on the results, the biggest advantage of the refactored code was the fact that it cut the run time to almost 4 times faster for each 2017 and 2018 analysis.  

### Pros & Cons of the original VBA Script
The pro of the original code was the fact that it was able to generated the same numbers and ultimately provide with the necessary analysis. However, the con is that the the refactored code cut so much time that it will be usable with a larger dataset. The original code was slower when it was ran and at times it was also laggy. 
