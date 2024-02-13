The VBA of Wall Street

![stock Market](Images/stockchart.jpg) 

### This project includes a smaller Excel workbook entitled "TestMacros" that allows a user to test the same modules from the "TyMultiple_year_stock_data" workbook on a smaller set of data. 

### Module 1 for the multiple year stock analysis worksheet generates total stock volumes for each individual stock on each individual worksheet. 

* This script will loop through each year of stock data and grab the total amount of volume each stock had over the year.

* It will display the ticker symbol to coincide with the total volume on each individual worksheet.


### Module2M includes additional column calculations as well as conditional coloring schemes based on the calculated columns. 

* This script will loop through all the stocks and take the following info.

  * Yearly change from what the stock opened the year at to what the closing price was.

  * The percent change from the what it opened the year at to what it closed.

  * The total Volume of the stock

  * Ticker symbol

* It is also conditionally formatted to highlight positive change in green and negative change in red.


### Module3H includes the maximum and minimum values for percent changes and the maximum total volume and along with the stock's ticker symbol. 

* This module includes everything from module 2. 

* Module 3 will also  locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" and puts them in a separate area on each worksheet starting in cell "O2".





