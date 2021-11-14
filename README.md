# Green-Stocks-Analysis
An overview of the performance of 12 selected green energy stocks over 2017 and 2018.

## Project Overview
The purpose of this project is to assist a recent Finance graduate (Steve) to analyze the performance of various clean energy stocks between 2017 and 2018. This is being done because Steve's first clients have a undiversified portfolio of just one clean energy stock, DAQO (DQ). As a financial professional, Steve knows that a diversified portfolio is a safer method of investing. So, he needs a way to carry out his analyses of 12 selected clean energy stocks. This project seeks to create a VBA script that can help him achieve this goal. The analysis for this project is based of two worksheets within the green_stocks.xlxs document. Within this document, stock metrics such as opening/closing price, high/low price, and volume are provided for 12 different stocks over the course of 2017 and 2018.

## Results
### Methods
Before analyzing the stock data, it is necessary that metrics of interest are identified so that code can be written to uncover insights. Given the data available, the most obvious metrics to find would be the rate of return for each stock for each year as well as the accompanying total volume of stock sold. So, a new sheet was created called "All Stocks Analysis". First, it needed to be determined which year needed to be analyzed. So, an input box was created which would create a variable (yearValue) to be used in code throughout the macro. Column headings for stock ticker, total daily volume, and return were created using the following code:

Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

Next, a string array, "tickers", was created to hold the ticker name's for each stock and the worksheet of interest was activated using the aforementioned yearValue variable. The row count of our data was also determined for later use in looping over the data:

Dim tickers(11) As String
tickers(0) = "STOCK1"
tickers(1) = "STOCK2"....

Worksheets(yearValue).Activate
RowCount = Cells(Rows.Count, "A").End(xlUp).RowRowCount = Cells(Rows.Count, "A").End(xlUp).Row

A ticker index was intialized at 0 to help refactor our earlier code and arrays of tickerVolumes, tickerStartingPrices, and tickerClosingPrices were created as Long, Single, and Single data types, respectively. A for loop was created to initialize the tickerVolume of each stock to zero. Another for loop was created to sequentially add the 'Volume' value of each row to tickerVolume(tickerIndex) until a new ticker appeared in column A. Once this occured, the value of the tickerIndex would be increased by 1. The first and last recorded 'Close' measurements for each ticker would also be stored in the startingPrices and endingPrices arrays, respectively. This was achieved with the following code:

For j = 0 To 11

tickerVolumes(j) = 0
    
Next j

    For i = 2 To RowCount
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    
    Next i

   The resulting populated arrays could then be used to fill in the required cells on the "All Stocks Analysis Sheet":
 
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1) = tickers(i)
        Cells(4 + i, 2) = tickerVolumes(i)
        Cells(4 + i, 3) = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
        
    Next i


Lastly, the "All Stocks Analysis" sheet needed to be formatted to reflect positive/negative return return rates as well as improving the aesthetic appearance of the headers on the worksheet. A timer also needed to be added using a startTime variable defined earlier in the code as well as an endTime variable as defined below:

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

In the end, the original (non-refactored) code ran took 0.707 seconds to run for 2017 and 0.705 seconds for 2018 (on attempts per year). Using the same measurement method, the refactored code ran in 0.118 seconds and 0.123 seconds, respectively. This means the refactoring of the original code saves approximately 0.6 seconds per run.
    
### Conclusions
From looking at the final results of the analyses, one can see that the stocks performed much poorer in 2018 than in 2017. In 2017, all stocks except for 'TERP' showed positive return rates. However, in 2018 the only stocks with positive return rates were 'ENPH' and 'RUN'. Since these two stocks showed only positive return rates, they could be considered more stable options to invest in as they have less of a risk of losing value based on our analysis.

However, the stocks 'ENPH', 'SEDG', and 'DQ' showed the highest overall return rates, in descending order, when both years were taken into account. This could indicate that 'SEDG' and 'DQ' may provide good returns, but could be more volatile than stable options like 'ENPH' and 'RUN'. Overall, 'ENPH' provided the best combination of stability and returns so, it would be wise for Steve's parents to divesify their portfolio through 'ENPH', 'RUN', and 'SEDG', with the bulk of their investments in 'ENPH'. 

## Summary
### Pros and Cons of Refactoring Code
In general, refactoring is something that should be done when you have the time, resources, and necessity to do so. Refactoring can help your code run faster and be more straight forward. This could help to prevent issues like long run-times and bugs whilst improving code maintainability in the future. It may seem that refactoring code is something that should always be done. This is not always the case, however.

While refactoring has all the benefits mentioned above, it costs time and money to do so. So, one should evaluate the cost to benefit ratio of refactoring their code before doing so. If the original code is stable and relatively simple then refactoring may be unecessary as you could be shortening the run time of your code by a very small amount whilst not making significant improvements to the readability of the code. This becomes more of a problem when taking into account that the refactoring process will be costly for the company who is using the code. It's not worth refactoring if the benefits do not outweigh the costs.

Another situation where refactoring may not be appropriate is when the dealine for a project is soon approaching as it may take longer than one would think and cause delays that cost one's employer money. Even if one can refactor the code in time, there may be bugs introduced to the code that will take even more time to fix. So, even if one's original code may be messy and unoptimized, one should evaluate whether the benefits of refactoring the code would outweigh the costs of delaying the project.

### Refactoring in This Analysis
The refactoring in this project was centered around creating arrays for ticker volume, starting prices, and ending prices instead of using 3 sererate for loops to achieve the same task for each metric. In this case, this decrased the run-time of the code by more than half a second. Obviously, this is a sizeable improvement over the original code. However, the original code was easier to understand as a beginner to VBA and coding in general. I believe that this is due to the fact that there are less variables in play in the original code as there is no tickerIndex, or arrays to be populated. I found it a struggle to understands the function of these objects until I went through the code step-by-step, taking the time to fully understand what values they contained and how they would be used in later steps of the code. When compared to the original code, it took very little time for me to understand how each for loop is storing and displaying values from the active worksheet. With this in mind, the only benefit of the original code is that it may be quicker to create and easier to maintain if the people working with the code are inexperienced in VBA or coding as future updates would be easier to implement and the time to create the code would be shorter. If the employer, Steve in this case, has people experienced in coding working for him then the refactored code is by far the better option as it can be created in roughly the same time and also result in greater efficiency.

