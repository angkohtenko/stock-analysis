# Stock Analysis with VBA. Code refactoring
## Overview of Project
I've got two datasets of stocks traded for 2017-2018 years. Both datasets are about 3000 rows. My goal was to write a VBA script to analyse trading volumes and annual return. After completing the macro, I refactored it and reduced execution time, so it can be easily used with bigger datasets.

## Results

### Creating pivot table
First of all, I wrote the macro to create a pivot table where we can find total trading volumes for each stock and the annual return. In addition, I created conditional formatting based on the annual return: green color for profitable stocks and red color for negative return.
For convenience I assigned macros to buttons and placed them on spreadsheet. It's easy to choose the year to analyse typing in pop-up window.

![Pop-up window](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/Pop-up%20window.png?raw=true)

![2017](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/2017.png?raw=true)

![2018](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/2018.png?raw=true)

As we can see, most of the stocks have negative return in 2018. So I’ll definitely pay my attention to ENPH. It brought more than 80% return to shareholders in 2018 and 129% in 2017.

### Code refactoring
I measured the efficiency of macro by setting a timer in the script. At the beginning I had this:
![Execution time for 2017](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/Execution%20time%20for%202017.png?raw=true)

![Execution time for 2018](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/Execution%20time%20for%202018.png?raw=true)

It may look good with current datasets, but I wasn't sure about larger datasets. I had a loop over all tickers and wrote a result in a pivot table for every ticker separately. My script was:

``` 
'Create a list of tickers
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
    
    
    Worksheets(yearValue).Activate
    
    'find the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through all tickers
    For i = 0 To 11
            ticker = tickers(i)
            
            'set initial volume to zero
            totalVolume = 0
            
            Worksheets(yearValue).Activate
            
            'Loop through rows in the data
            For R = rowStart To rowEnd
            
                    'Increase totalVolume
                    If Cells(R, 1).Value = ticker Then
                         totalVolume = totalVolume + Cells(R, 8).Value
                     End If
                     
                    'set starting price
                    If Cells(R, 1).Value = ticker And Cells(R - 1, 1).Value <> ticker Then
                         startingPrice = Cells(R, 6).Value
                     End If
                     
                     'set closing price
                     If Cells(R, 1).Value = ticker And Cells(R + 1, 1).Value <> ticker Then
                         endingPrice = Cells(R, 6).Value
                     End If
                     
            
            Next R
            
            Worksheets("All Stocks Analysis").Activate
    
            Cells(i + 3, 1).Value = ticker
            Cells(i + 3, 2).Value = totalVolume
            Cells(i + 3, 3).Value = endingPrice / startingPrice - 1

   Next i 
   ```
After refactoring I created three output arrays: 

```
Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
```

and made a loop to save values to these arrays at first and print arrays into the pivot table in the end.

```
''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        '3d Increase the tickerIndex.
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
By this little trick execution time was reduced by 80%.

![Refactored Execution time for 2017](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/Refactored%20Execution%20time%20for%202017.png?raw=true)

![Refactored Execution time for 2017](https://github.com/angkohtenko/stock-analysis/blob/main/Resources/Refactored%20Execution%20time%20for%202018.png?raw=true)


## Summary
The major benefits of code refactoring are reducing execution time and improving a readability. However, it requires a time to review the code and, which is more important, it’s hard to find an elegant solution right away. Usually, I prefer to refactor code in a while, when I have a fresh eye.

In this case original script was a straightforward, but inefficient: loop over all rows, switch constantly between worksheets and print value by value.
Refactored script isn’t complicated at all, but more efficient. It’s important to know what kind of operations are resource-consuming. During refactoring we can avoid them and use high-level solutions.

