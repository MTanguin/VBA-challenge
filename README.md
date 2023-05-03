# VBA-challenge
VBA Functions | Iterations | Conditional Formatting | Working with Worksheets


Credit:
E. Accime.


This application is a useful analysis of stock market data that can help investors make informed decisions.

- It performs an analysis of the stock market data. It calculates the yearly change and percent change for each stock ticker, along with the total stock volume. It also identifies the ticker symbols with the greatest percent increase, greatest percent decrease, and greatest total volume. Finally, it highlights positive changes with green and negative changes with red.

- The code uses a loop to iterate through the rows of data and compare the ticker symbol of each row with the previous row. If the ticker symbol changes, the code calculates and prints the yearly change, percent change, and total stock volume for the previous ticker. If the total stock volume for the ticker is zero, the code prints zeros for the yearly change and percent change.

- The code also calculates the maximum and minimum percent changes and the maximum total volume, and prints the corresponding ticker symbols. It uses the MATCH function to find the row number of the ticker symbols with the maximum and minimum percent changes and maximum total volume.
