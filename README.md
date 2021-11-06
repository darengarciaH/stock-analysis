# stock-analysis

## Overview
Steve, a recent finance graduate, needed to further analyze green energy stocks for his parents, who are his first clients. First, they wanted to invest all their money in DQ, but then figured that they may need to diversify their portfolio. As a result, Steve sought our help in running Excel VBA code in order to analyze stock data in a short period of time. While we were able to write a successful VBA macro that can analyze data from a dozen stocks, our code may take too long to analyze thousands of stocks. Therefore, in this analysis, we will refactor our VBA code so Steve can use it to not only analyze data for a dozen stocks, but rather for thousands of them in a short period of time.

## Results
### Stock Analysis
Generally speaking, stocks in this dataset performed considerably better in 2017 than in 2018. Only ENPH and RUN had positive returns in both years, while TERP had negative returns for both years. All other stocks had positive returns for 2017 and negative returns for 2018, including DQ, where Steve's parents originally wanted to invest all their money in. 

<img width="323" alt="2017_StockResults" src="https://user-images.githubusercontent.com/92702922/140062180-e8c04690-fb3a-4dd3-998b-1818c7451de4.png">

<img width="322" alt="2018_StockResults" src="https://user-images.githubusercontent.com/92702922/140062224-4ea88dce-ffe8-4ff1-b841-4b9a76f6ff0b.png">

In 2017, several stocks reported returns of more than 100%, including DQ (199.4%), SEDG (184.5%), ENPH (129.5%), and FSLR (101.3%). All other stocks reported positive returns, with the exception of TERP with a -7.2% return. In 2018, however, only ENPH and RUN report positive returns of 81.9% and 84%. DQ had negative returns of -62.6%, being the worst performing stock in 2018 despite having been the best performing stock in 2017. 

### VBA Code
For this analysis, we have subroutines for both the original script and the refactored script.  The original script involved creating variables for the starting price, ending price, and total volume that were rewritten every time a different ticker was read into our for-loop. See the example below from the original script:

```
For i = 0 To 11
    ticker = tickers(i) 'loads all tickers one by one, then checks row by row
    totalVolume = 0
    
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        If Cells(j, 1).Value = ticker Then
            'adds volume when ticker matches
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
        'checks if previous row ticker and current row ticker are different to use starting price
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If
        'checks if current row ticker and one after are different to denote ending price
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        End If
    Next j
```
The key difference that we implemented for the refactored script involved creating arrays for the ticker starting price, ticker ending price, and ticker total volume, which included such values for each of the 12 tickers. Here, we are not rewriting variables that already had values stored for each of the starting prices, ending prices, and total volumes, but rather we stored them into their individual index in each array. See the code below:

```
For tickerIndex = 0 To 11
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            '3a) Increase volume for current ticker
            If Cells(i, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            '3c) check if the current row is the last row with the selected ticker
            ' If the next row’s ticker doesn’t match, increase the tickerIndex
             If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            '3d Increase the tickerIndex
        Next i
    Next tickerIndex
```

Because of these changes, we can see that this is probably the reason as to why the refactored code takes less time to run. Here are the original script timers:

<img width="259" alt="OriginalStocks_2017" src="https://user-images.githubusercontent.com/92702922/140062751-81dc5716-9574-46b7-89b2-13adfe44b9ef.png">

<img width="252" alt="OriginalStocks_2018" src="https://user-images.githubusercontent.com/92702922/140062780-a12c8723-0c83-43c8-90bc-e78a8d0d7119.png">

The refactored script takes about three seconds less to process the same results for the stock analysis, as can be seen below:

<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92702922/140062975-4e232795-a893-4fdf-a065-14d6ba0f5f1d.png">

<img width="254" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92702922/140062989-5b15d36a-9c51-4bf2-b34c-cf75232d1292.png">

Since the refactored script can be much more easily applied to datasets that involves thousands of stocks, we can see that the faster processing time is good news for Steve's future stock analyses.

## Summary
The advantages of refactoring code make programs run faster and more efficiently. This can make code less concrete and more versatile to be applied to larger numbers of stocks (as in this analysis), and as a result, can be used in different scenarios as well. The disadvantages might include coding to be more detail-oriented, making it prone for creating mistakes that will allow code to not run. It would also be difficult for time-sensitive projects since refactoring can take extra time to complete, while also running the risk of introducing errors in code that would prolong the length of these projects even more.

For this example, the advantages involve being able to process larger datasets that contain data for thousands of stocks in a shorter amount of time, and not just a dozen changes. This would be great for Steve for future projects where he would have to analyze thousands of stocks for his clients. However, the disadvantage of refactoring would be that Steve would need to spend more time to do this, and would be detrimental for time-sensitive projects. Steve would also run the risk of modifying his code and encountering bugs that might take longer to fix or make his macros non-functional until potential bugs are fixed. Ultimately, the advantages and disadvantages of refactoring code would vary depending on Steve's timeframe to complete his projects: if there is sufficient time, refactoring will make his code easier to read and apply for future and similar projects and would be a worthy time investment. If time is lacking and deadlines are impending, it might not be worth taking the extra time especially if his program is perfectly functional. 
