# stock-analysis

### Purpose of this project is to analyze Green Enery stock data for years 2017 - 2018 and provide insights on which is the best to add to the portfolio. VBA has been used to acheive this. The VBA script has also been refactored to make the macro more robust and faster. Details are provided below.

## Analysis of Green Energy Stocks

### Analysis of DAQO New Energy Corp (Ticker: DQ)

  - After carefully analyzing and reviewing the stock performance of DQ for the year 2018, it has been observed that this ticker has not given any positive returns and has resulted in a negative return of -62.6%. Due to this, it is highly recommeneded to avoid adding this company to your portfolio.

##Results:

### Analysis of all other for years 2017 - 2018.

  - At first, the output sheet is formatted by adding a Sheet header and row headers. This is done by using the below code:
    ```
        Format the output sheet on the "All Stocks Analysis" worksheet.
        'Sheet Header
        Cells(1, 1) = "All Stocks (" + yearvalue + ")"
        
        'Create a header row
        Cells(3, 1) = "Ticker"
        Cells(3, 2) = "Total Daily Volume"
        Cells(3, 3) = "Return"
    ```
    
    
