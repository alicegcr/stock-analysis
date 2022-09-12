# Stock-analysis

## Purpose
The motivation behind of this VBA challenge script is to loop through green energy stock data to analyze, and document meaningful information for each stock. The script analysis results on the stocks that had the best performance, in profit and total volume.

## Assumptions
- each worksheet represents a different year 
- the tikcers are organized by alphabetized 
- there are 7 columns: ticker, date, open price, high price, low price, close price, volume 

## Modules
This VBA module loops through all sheets in the workbook:
1. ticker symbol
2. the yearly change from opening price at the beginning of a given year to the closing price at the end of that year. The cells are conditionally formatted such that a cell with a zero/positive yearly change is green and a cell withe a negative yearly change is red
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year. Cells are formatted with percent format
4. The total volume or the stock for that year

