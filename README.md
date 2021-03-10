# Stocks-Analysis

## Overview
Using VBA, this project provides a full spread of core programming concepts while organizing information relating to stock data.
With only minor adjustments this code will surely be something that anyone working in a financial field would use for the entirety of their career.

## Results
Below are images of two very similar programs processing data from two worksheets named "2017" and "2018".
These worksheets contain the daily share information for different companies during the year the worksheet is named after: 
  Ticker, Date, Open, High, Low, Close, Adj Close, and Volume

For the sake of this project, we are focusing on the amount that each company's stock price changed from year start to year end, as well as the total volume of shares that changed hands during that time period.


![gfx1](https://user-images.githubusercontent.com/14188580/110701270-95ab6200-81b6-11eb-9aba-f26cd1628ce1.png)

![gfx2](https://user-images.githubusercontent.com/14188580/110701277-96dc8f00-81b6-11eb-827f-d52b338aeb16.png)

While the outputs are nearly identical, there is a significant difference in their processing time.

The macro on the left uses a nested For loop to iterate between every row of data (j) within iterating for every ticker name (i).
Before each iteration of our upper loop, we are outputting the ticker name, total volume of all trading days, and the difference in price into our destination worksheet - before resetting our variables to 0 and getting dizzy from another loop for a new ticker symbol.
Each worksheet in this example has around 3000 rows of data, multiply that by the 12 tickers we have and we have a program that is performing over 36,000 iterating processes, taking almost 0.7 seconds to complete.

While we currently live in 2021 and today's cutting edge tech doesn't even blink at these numbers, it is always prudent to write the most efficient and legible code possible.

The images on the right show the output of a macro that refactors our original idea with a little more ~pizazz~.
By utilizing Arrays in VBA, we can use a single pass through all of our ~3000 rows to "append" information to them by an increasing index for each ticker.
At the end of this pass, we can output the information we've stored with a short loop, cutting out ~0.5seconds of processing time.

The only negative aspect of this refactoring is the higher likelyhood of the code breaking when bring in new datasets:
  * The "If Then" system that we use while iterating through the rows relies on our column data being sorted by ticker.
  * If all of one ticker's information is not sequential on the worksheet, the macro will almost certainly break by incrementing our index beyond the Array length.
  * Having clean data to begin with is always a plus!

## Summary
Refactoring code can greatly increase your productivity and efficiency.
One possible drawback of refactoring could be trading legibility for efficiency; and in a professional setting, legibility can be just as important as efficiency when working with others. By refactoring this project with several Arrays the code becomes very chewy compared to how simple the original process's logic is.

As with anything, proper documentation and comments can go a long way!

