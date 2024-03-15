# VBA-Challenge
Repo for Module 2 Challenge

# Our challenge required us to create a VBA script to loop through all the stocks for a given year and output the following:
- Ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year 
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
- The total stock volume of the stock to match a picture provided in the challenge

####################################################
- Screenshots are named after year tabs in workbook
- alphabetical_testing_JL.xlsm was used for testing code 
- Multiple_year_stock_data_JL.xlsm is the completed workbook
- Module 2 Challenge Stock Analysis.vbs is the separated VBA script 

# I used Dim to assign variables as values 

# I used a For Loop to loop through each worksheet

# Defined ranges and input column header info for the output

# Assigned variables for additional functionality and set starting values for each

# Used a For Loop to loop through all rows

# Checked for Ticker, EndPrice, YearlyChange, PercentChange, TotalVolume values for summary table

# Assigned value to ticker, EndPrice

# Calculated YearlyChange, PercentChange, TotalVolume

# Created summary table with output values

# Used conditional formatting to fill color based on cell value

# Checked for greatest increase, decrease and volume using an If statement

# Checked for Ticker with highest volume

# Moved to next row and reset value

# Increase TotalVolue by value in Volume (i,7)

# Set the start price for the next stock

# Set startprice to Value of Open (i,3)

# Output the greatest values and their corresponding tickers