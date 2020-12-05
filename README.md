# Stock Analysis

## Overview of Project
The purpose of this project was to refactor VBA code to decrease the run-time for analyzing a set of stocks. In a previous analysis ([green_stocks.xlsm](https://github.com/borkard/stock-analysis/blob/main/green_stocks.xlsm)) I had written and run a VBA code to analyze the total daily volume and return for a dozen stocks, but wanted to expand the analysis to include the entire stock market. The code for the previous analysis took a long time to run and would take even longer if it had to run through and analyze many more stock tickers. In the new analysis, the VBA_Challenge below, I tried to improve the efficiency of the code by refactoring it and using fewer steps including creating a ticker index and removing nested For loops.

**VBA Stock Analysis Excel File:** [VBA_Challenge.xlsm](https://github.com/borkard/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Results
Comparing the analysis of all the stocks in 2017 and 2018, nearly all stocks fared better in 2017. The stocks with the greatest return in 2017 were DQ, SEDG, ENPH, and FSLR; all with well over a 100% return rate. Only one stock, TERP, had a negative return in 2017. By contrast in 2018, nearly all stocks had a negative rate of return. With the exception of ENPH, all of the stocks that had the highest rate of return in 2017 had a negative return in 2018. Only two 2018 stocks had a positive rate of return: ENPH(81.9%) and RUN(84%).


The images below show the post-refactoring run time for the analysis of 2017 and 2018 stocks. In the initial analysis, the run times were 0.953125 seconds and 0.9296875 seconds respectively. Refactoring the code greatly improved the efficiency of the analysis and run times as it does not have as many steps to run through and is laid out logically. 

![VBA_Challenge_2017](https://github.com/borkard/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018](https://github.com/borkard/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

Where nearly four functions were all nested in one For loop in the initial analysis, breaking them up into four separate For loops in the refactored code helped to process the analysis faster. A snippet of the refactored code is below:

''2b) Loop over all the rows in the spreadsheet.

        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
                startingPrice = Cells(j, 6).Value
                tickerStartingPrices(tickerIndex) = startingPrice
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                endingPrice = Cells(j, 6).Value
                
                tickerEndingPrices(tickerIndex) = endingPrice
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            'End If
            End If
            
        Next j

Creating arrays and a ticker index also helped to organize and store the information more logically. A snippet of the refactored code is below:

   '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Double
    Dim tickerEndingPrices(12) As Double   


## Summary
Refactoring code is very useful to improve the efficiency and logic of the code. Refactoring code can also help you find errors or duplicative lines of code. It is also helpful to refactor code if others will also be working on it so that the code is clean and others can follow along. Despite these benefits of refactoring code, it is also time consuming to find bugs and alternative methods for the code. Another disadvantage is that consistency in the formatting of the code may be lost. When refactoring the original VBA script, I personally spent a lot of time working through errors in different lines and getting lost in what was the original code and what I was intending to do. It also took some time to figure out what worked to make it run faster, such as splitting up the For loops rather than having one large nested For loop, but I am glad that it now runs efficiently and that I can easily understand my own code with the comments and be able to share it with others.

