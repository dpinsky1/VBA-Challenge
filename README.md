# VBA-Challenge

Many thanks to Saad Khan and Angel Milla for editing and guidance on the code used in this challenge. 

Documentation used as reference and foundational code:

Match functionality
https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.match

Cell references for ticker
https://github.com/emmanuelmartinezs/stock-analysis

Saad Khan: editing "for loops" and conditionals
Angel Milla: guidance on match and printing of data analysis 

This code will loop through all sheets in a workbook and pull unique values from the first column. It will then pring those values in column "I". 

Assuming all like values are grouped together (as these are sorted by date), upon finding the next unique value, the script will then take the first open, and the last close, and perform calculations to find the total change for the specified timeframe, and the percent change over that same timeframe. 

Next the script will tabulate all volumes (column "G")
and display the total next to the appropriate unique value in column "I".

Conditional formatting is added to illustrate the positive and negative growth in column "J". 

Using this compiled data, a summary is printed in range "O1:Q4".
This summary includes the greatest percentage increase, greatest percentage decrease, and the greatest total volume for the tickers for the given years. 