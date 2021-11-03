# stock-analysis

## Overview
Steve, a recent finance graduate, needed to further analyze green energy stocks for his parents, who are his first clients. First, they wanted to invest all their money in DQ, but then figured that they may need to diversify their portfolio. As a result, Steve sought our help in running Excel VBA code in order to analyze stock data in a short period of time. While we were able to write a successful VBA macro that can analyze data from a dozen stocks, our code may take too long to analyze thousands of stocks. Therefore, in this analysis, we will refactor our VBA code so Steve can use it to not only analyze data for a dozen stocks, but rather for thousands of them in a short period of time.

## Results
### Stock Analysis
Generally speaking, stocks in this dataset performed considerably better in 2017 than in 2018.

<img width="323" alt="2017_StockResults" src="https://user-images.githubusercontent.com/92702922/140062180-e8c04690-fb3a-4dd3-998b-1818c7451de4.png">

<img width="322" alt="2018_StockResults" src="https://user-images.githubusercontent.com/92702922/140062224-4ea88dce-ffe8-4ff1-b841-4b9a76f6ff0b.png">

Only ENPH and RUN had positive returns in both years, while TERP had negative returns for both years. All other stocks had positive returns for 2017 and negative returns for 2018, including DQ, where Steve's parents originally wanted to invest all their money in. 

### VBA Code
The original script takes longer to process than the refactored code. Here are the original script timers:

<img width="259" alt="OriginalStocks_2017" src="https://user-images.githubusercontent.com/92702922/140062751-81dc5716-9574-46b7-89b2-13adfe44b9ef.png">

<img width="252" alt="OriginalStocks_2018" src="https://user-images.githubusercontent.com/92702922/140062780-a12c8723-0c83-43c8-90bc-e78a8d0d7119.png">

The refactored script takes about three seconds less to process the same results for the stock analysis, as can be seen below:

<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92702922/140062975-4e232795-a893-4fdf-a065-14d6ba0f5f1d.png">

<img width="254" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92702922/140062989-5b15d36a-9c51-4bf2-b34c-cf75232d1292.png">

Since the refactored script can be much more easily applied to datasets that involves thousands of stocks, we can see that the faster processing time is good news for Steve's future stock analyses.

## Summary
The advantages of refactoring code make programs run faster and more efficiently. This can make code less concrete and more versatile to be applied to larger numbers of stocks (as in this analysis), and as a result, can be used in different scenarios as well. The disadvantages might include coding to be more detail-oriented, making it prone for creating mistakes that will allow code to not run. It would also require formatting our code more and adding more comments and descriptions so that code can be easier to understand for others.

For this example, the advantages involve being able to process larger datasets that contain data for thousands of stocks in a shorter amount of time. The disadvantages would still inco
