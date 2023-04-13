An Excel workbook containing multiple worksheets of stock market data is analyzed by a VBA macro that performs the following tasks:

1. Iterates through each worksheet and activates it.
2. Determines the range of data in column A of the current worksheet and stores it as a string variable.
3. Extracts the unique stock ticker symbols from the determined range of data and saves them to an array.
4. For each unique stock ticker symbol, identifies the first and last dates associated with it, and calculates the open and close values as well as total volume.
5. Calculates the yearly and percent changes in value for each ticker symbol.
6. Identifies the ticker symbol with the greatest percent increase, greatest percent decrease, and greatest total volume, and saves the corresponding values and ticker symbols to cells in the worksheet.
7. Colors the cells in column L based on whether the percent change for each ticker symbol is positive, negative, or zero.