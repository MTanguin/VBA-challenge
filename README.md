#### "Using VBA in analysing stock market data"
VBA Functions | Iterations | Conditional Formatting | Working with Worksheets


Credit:
E. Accime.

# Overview
This application is a useful analysis of stock market data that can help investors make informed decisions.

- It performs an analysis of the stock market data. It calculates the yearly change and percent change for each stock ticker, along with the total stock volume. It also identifies the ticker symbols with the greatest percent increase, greatest percent decrease, and greatest total volume. Finally, it highlights positive changes with green and negative changes with red.

- The code uses a loop to iterate through the rows of data and compare the ticker symbol of each row with the previous row. If the ticker symbol changes, the code calculates and prints the yearly change, percent change, and total stock volume for the previous ticker. If the total stock volume for the ticker is zero, the code prints zeros for the yearly change and percent change.

- The code also calculates the maximum and minimum percent changes and the maximum total volume, and prints the corresponding ticker symbols. It uses the MATCH function to find the row number of the ticker symbols with the maximum and minimum percent changes and maximum total volume.

# Methods

Created a script that loops through all the stocks for one year and outputs the following information:

- The ticker symbol

- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

- The total stock volume of the stock. 

- Added functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

- Made the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

# Results

![2018 StockMarketAnalysis](https://github.com/MTanguin/VBA-challenge/assets/114210481/7377c44e-f8c1-4fc0-a79e-787986aa2312)


Source:

https://courses.bootcampspot.com/courses/2799/assignments/42915?module_item_id=802686
