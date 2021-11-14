# Green-Stocks-Analysis
An overview of the performance of 12 selected green energy stocks over 2017 and 2018.

##Project Overview
The purpose of this project is to assist a recent Finance graduate (Steve) to analyze the performance of various clean energy stocks between 2017 and 2018. This is being done because Steve's first clients have a undiversified portfolio of just one clean energy stock, DAQO (DQ). As a financial professional, Steve knows that a diversified portfolio is a safer method of investing. So, he needs a way to carry out his analyses of 12 selected clean energy stocks. This project seeks to create a VBA script that can help him achieve this goal. The analysis for this project is based of two worksheets within the green_stocks.xlxs document. Within this document, stock metrics such as opening/closing price, high/low price, and volume are provided for 12 different stocks over the course of 2017 and 2018.

##Results
###Methods
Before analyzing the stock data, it is necessary that metrics of interest are identified so that code can be written to uncover insights. Given the data available, the most obvious metrics to find would be the rate of return for each stock for each year as well as the accompanying total volume of stock sold. So, a new sheet was created called "All Stocks Analysis". First, it needed to be determined which year needed to be analyzed. So, an input box was created which would create a variable (yearValue) to be used in code throughout the macro. Column headings for stock ticker, total daily volume, and return were created using the following code:

Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

Next, a string array, "tickers", was created to hold the ticker name's for each stock and the worksheet of interest was activated using the aforementioned yearValue variable. The row count of our data was also determined for later use in looping over the data.

Dim tickers(11) As String
tickers(0) = "STOCK1"
tickers(1) = "STOCK2"....

Worksheets(yearValue).Activate
RowCount = Cells(Rows.Count, "A").End(xlUp).RowRowCount = Cells(Rows.Count, "A").End(xlUp).Row

A ticker index was intialized at 0 to help refactor our earlier code and arrays of tickerVolumes, tickerStartingPrices, and tickerClosingPrices were created as Long, Single, and Single data types, respectively. A for loop was created to initialize the tickerVolume of each stock to zero. Another for loop was created to sequentially add the 'Volume' value of each row to tickerVolume(tickerIndex) until a new ticker appeared in column A. Once this occured, the value of the tickerIndex would be increased by 1. The first and last recorded 'Close' measurements for each ticker would also be stored as startingPrice and endingPrice, respectively. This was achieved with the following code:

For i = 0 To 11
        tickerVolumes(i) = 0
        
For i = 2 To RowCount
   
   tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
      
     If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then:
            tickerStartingPrices(i) = Cells(i, 6).Value
        
        
     End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then:
           tickerEndingPrices(i) = Cells(i, 6).Value
            tickerIndex = i + 1
            
        End If
            
    
    Next i
    
    
