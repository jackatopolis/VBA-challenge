### VBA-challenge
### By: Jack Cohen

# PURPOSE
The purpose of this VBA script is to loop through stock data to extract, analyze, and document meaningful data for each stock in each year of data provided.

# ASSUMPTIONS
The module was created to analyze stock data with the following assumptions:
1. Each worksheet represents a different year
2. Tickers in each sheet are alphabetized, not randomly ordered
3. Columns for raw data are, in order: ticker, date, open price, high price, low price, close price, volume

# MODULES
This VBA module loops through all sheets in a workbook. In each worksheet, the module loops through all stocks in a year (each worksheet represents one year) and outputs:
1. The ticker symbol
2. The yearly change from opening price at the beginning of a given year to the closing price at the end of that year. The cells are conditionally formatted such that a cell with a zero/positive yearly change is green and a cell withe a negative yearly change is red
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year. Cells are formatted with percent format
4. The total volume or the stock for that year