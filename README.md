# Stock Market Analysis
Using VBA to analyze stock market data.

## File Description
This VBA script was tested on multiple year stock data. The file consisted on three tabs. Each tab consisted of: 

    Column A: Stock Ticker
    Column B: Date
    Column C: Opening stock value
    Column D: Highest stock value
    Column E: Lowest stock value
    Column F: Closing stock value
    Column G: Total stock volume
   
## VBA Description
The VBA script loops through all stocks on each tab and creates a table on each worksheet that calculates yearly 
change, percent change, and total stock volume by ticker. If yearly change was greater than 0, the cell is green. If it 
was less than 0, the cell is red.

    Yearly Change = Total closing stock value - Total openning stock value
    Percent Change = (Total closing stock value - Total openning stock value)/Total openning stock value

You only need to run the script once to loop through all the tabs.
