# Stock Summarization Using VBA in Excel

![Stocks](images/stocks.jpg)

## Background

In this repository I used VBA scripting to analyze real stock market data in CSV format. These CSV files contained three years of stock data and more than 2 million rows.

## Analysis Summary

* Created a script that loops through all the stocks for one year for each run and takes the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Created conditional formatting that will highlight positive change in green and negative change in red.

* Returned the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

* Made appropriate adjustments to my VBA script that allows it to run on every worksheet, i.e., every year, just by running the VBA script once.
