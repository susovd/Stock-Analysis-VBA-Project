<!---Project Logo -->
<br />
<p align="center">
  <a href=>
    <img src="Images/stockmarket.jpg">
  </a>

  <h3 align="center">World Demographic Trends 1950-2019</h3>
  <p align="center">
    A Data Visualisation Project
    <br />
</p>
</p>


<!-- ABOUT THE PROJECT -->
## About The Project

## Background

In this project I used VBA scripting to analyze real stock market data. 

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - I used this file to develop my scripts. This data set is smaller and allows you to test faster. This code should run on this file in less than 3-5 minutes.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - I ran my scripts on this data to generate the final report.


I created a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* I also have conditional formatting that will highlight positive change in green and negative change in red.


My solution also returns the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 

The VBA script runs on every worksheet, i.e., every year, just by running the VBA script once.


**Additional reference materials:**

_Best-README-Template_ Retrieved from: [https://github.com/othneildrew/Best-README-Template](https://github.com/othneildrew/Best-README-Template)
