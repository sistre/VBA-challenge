# VBA-challenge
Data Bootcamp VBA-challenge Homework

# Author
Sean Istre 
[sistre@mac.com](mailto:sistre@mac.com)

# What this Visual Basic Script Accomplishes
The Stock_Analysis Visual Basic Script is designed to analize a year's series of daily stock reports determining the Yearly Change, Percent Change, and Total Stock Volume (reported in thousands due to date size limitations). Then using this data it finds the stocks with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume (in thousands).

# Worksheet Starting Layout Requirements
This script requires that the data for each day be entered into specific cells in the following order sorted by ticker:
Column | Header | Data
--- | --- | ---
A | &lt;ticker&gt;	| Stock Ticker
B | &lt;date&gt; | Unused
C | &lt;open&gt; | Opening Stock Price
D | &lt;high&gt; | Unused
E | &lt;low&gt; | Unused
F | &lt;close&gt;	| Closing Stock Price
G | &lt;vol&gt; | Stock Volume

# Output
The script analyzes the data in the daily report data and outputs the results in the following columns:
Column | Header | Data
--- | --- | ---
I | Ticker	| Stock Ticker
J | Yearly Change | Difference of the stock price at the beginning of the year and the end 
K | Percent Change | The percent change of the stock for the year
L | Total Stock Volume (K) | The total stock volume for the entire year in thousands

Using this data it outputs the Ticker in column 'P' and the Value in column 'Q' for the following rows
Row | Calculation
--- | ---
2 | Greatest % Increase
3 | Greatest % Decrease
4 | Greatest Total Volume (K)

# Notes
Percent Change cannot be calculated for stocks that have and open stock price of 0 for the beginning of the year. The Percent Change for these stocks is denoted at Not a Number (NaN).
