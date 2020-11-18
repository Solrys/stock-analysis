# Stock Analysis
## Project Overview

A friend of mine named Steve asked me to analyze a set of stocks with an excel worksheet. Steve wanted to be able to visualize the performance of these set of stocks so that he could inform his parents on good choices for investment within the group. Initially I created a workbook to provide Steve with the data he requested so that at a touch of a button he could visualize and compare stocks based on their volume and return. 
Steve liked the workbook I prepared for him. Now Steve wants to be able to use this workbook for an expanded dataset to include the entire stock market over the last few years. Because Steve will be comparing a much larger amount of stocks, refactoring the VBA code becomes necessary. Refactoring the code will make the code more efficient by taking fewer steps, using less memory, and in this case hopefully creating shorter Run Times for Steve.


## Results

In order to provide Steve with his requested criteria, I needed to change the code and refactor it, making it more efficient. I did this by changing the nesting order and removing an additional variable (j). In this new workbook I introduce the tickerIndex and set it to zero. I then created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The “tickers” array was used to establish the ticker symbol of a stock. These three arrays corresponded with the tickers array by using the initial tickerIndex variable I set up. Being able to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before running the loops through the data set allowed for faster run times. 

### A Closer Look At The Diferentiation in VBA code

#### Original Code:
    '3b) Activate data worksheet
   
    Worksheets(yearValue).Activate

    '3c) Get the number of rows to loop over
   
     RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers
   
    For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate
    
       '5) loop through rows in the data
       
    For j = 2 To RowCount
    
           '5a) Get total volume for current ticker

     If Cells(j, 2).Value = ticker Then

            'increase totalVolume by the value in the current row
            TotalVolume = TotalVolume + Cells(j, 9).Value
    
    End If
    
           '5b) get starting price for current ticker

        If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set starting price
            startingPrice = Cells(j, 7).Value

        End If

           '5c) get ending price for current ticker
           
           If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set ending price
            endingPrice = Cells(j, 7).Value



#### Refactored Code:
       'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays.
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
     '2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
       '3a) Increase volume for current ticker 'Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex
        If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
     tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
            

            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
For The Full VBA Code see the following link to the file VBA_Challenge.vbs
https://github.com/Solrys/stock-analysis/blob/main/VBA_Challenge.xlsm



Once I established a refactored code, I ran the original code for 2017 and 2018, and compared the Original Run Times with the Refactored Run Times.
#### Refactored Run Times:
![refactored 2017](https://github.com/Solrys/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![Refactored 2018](https://github.com/Solrys/stock-analysis/blob/main/Resources/Resources/VBA_Challenge_2018.png)

After comparing both sets of runtimes it was determined the the new Refactored code was more efficient and ran through all 12 stock tickers at about .5 seconds faster!

## Summary

The Advantage of a refactored code is creating a more efficient code, taking fewer steps, making the code more direct, and potentially saving time with "for loops" to allow for larger datasets. The possible disadvantages of refactoring code is that there lies a possibilty of creating a mistake in the code while trying to refactor. Omitting a variable or a step may render the data innacurate. 

### Advantages and Disadvantages of Original VBA Script
In this specific case the advantage of the Original VBA script was that the code was easier to arrive at. having a nested loop witha second variable kept the two variables easy to distinguish when writing the code. (i,j) The disadvantage of the Original code is that it took more time to run, and with a significantly larger dataset the delays may be more substantial. 

### Advantages and Disadvantages of Refactored VBA Script
The main advantage of the Refactored VBA script was that it saved run time and proved more efficient. This code will work better with a larger dataset. Using VBA also allows you to write both codes side by side while testing the efficiency using the play feature on the ribbon. The disadvantage of Refactored VBA script is that it may introduce new bugs, and sourcing the percise methods for making the code refactored may be complex and may potentially cause new paths for error.  
