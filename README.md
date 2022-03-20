# Stock-analysis

## Overview and purpose of the project
We are doing this project to help Steve working in the financial industry find accurate investment
information that could be useful to his parents who need to invest in green energy. As they do not 
have sufficient knowledge about investment, they chose to invest all their money into Daqo New 
Energy Corporation is a company that makes silicon wafers for solar panels. However, Steve would prefer 
diversifying his parent's funds and decided to analyze multiple green energy stocks in addition to Daqo's
stock using an excel file. To assist Steve, instead of using excel functions directly to run the analysis, 
we will use VBA built to automate tasks and employ complex logic. Also, this will permit Steve reuse the 
code we wrote with any stock while reducing the chance of errors.


## Analysis

# Total Volume
The Data set we analyzed contains information about a dozen stocks. We measured how actively 
a stock is traded by computing the total number of shares traded over the year. To proceed, we used the 
information in the column "Volume" in the file "green_stock.xlsm" which gives the number of shares traded 
by each stock in a day. Since the data in the sheet is sorted by stock and by date, each row containing 
information for one stock for a specific day, to compute the total volume, in a VBA subroutine, we loop over 
the rows retrieving the volume traded daily for a stock and adding this number to the value of the previous day 
saved in a variable of type double. 


# Yearly return
The second piece of information that we decided to compute from the data set is the yearly return for each stock 
which is the percentage difference in price from the beginning to the end of the year. To proceed, 
within the same Subroutine, while retrieving the daily volume for a stock in a row, we also check if this given 
row contains the price for the beginning of the year or the price for the end of the year and saves those values 
in two variables of type double. 


Once we have the total volume, the starting and ending price for a stock, before computing the same values for 
the following stock and overwrite our variables, we insert them in the "DQ Analysis" worksheet created for the results.
 

# Formatting the results
To make it easier for Steve to read the results, we added code to our subroutine to format the sheet "DQ Analysis" 
To help him determine at a glance which stocks performed well and which ones did not, we format the cells based on 
the values of the returns, making positive returns green and negative returns red using conditionals. Also, as 
Steve needed a way to run these analyses quickly, we created a button and linked it to our subroutine.

As in the data set, we have information for the stocks for two years, 2017-2018; we want to give Steve the possibility 
to run an analysis for both. To make the macro interactive, we created an input box where Steve can enter the year 
for which he wants to run the analysis. To use the same subroutine to run the analysis for both years, we save the user 
input in a variable of type single. We use it to select the sheet to activate when we start retrieving the values for 
the total volume, the starting and ending price.

![2017](https://github.com/valerielnd/stock-analysis/blob/main/Analysis_Results_2017.png)

![2018](https://github.com/valerielnd/stock-analysis/blob/main/Analysis_Results_2018.png)

# Effiency of the code
The subroutine we wrote works well to analyze the data set for 12 stocks. However, in the future, Steve might want 
to run the same analysis on larger data sets. To check the efficiency of the subroutine, we added a script that will 
calculate how long the code takes to execute and output the elapsed time in a message box. The analysis for the year 2017 
was executed in 0.97 seconds and for the year 2018 in 0.99 seconds.

![original_2017]()

![original_2018]()

As we suspected, the subroutine is efficient for 12 data stocks, but to expand the dataset to include the entire stock 
market over the last few years, we will need to refactor the code in our subroutine.

# Refactor the code
The first change we made was to use three arrays of the same size of the number of stocks, one to keep their total volume, 
one their starting price, and the last one their ending prices. This modification will let us gather the values that we need for our 
analysis without going from the sheet that has the data to the sheet where we keep the results multiple times. 
To iterate over the arrays, we will use an index variable that gets incremented each time we are done with one stock and 
start working with the next.

Once we had all the values for each stock, we looped over the arrays and inserted the values in them in the results sheet. 
We kept the same button to run the analysis and used the same mechanism to calculate how long the code took to perform the analysis. 
With those modifications, the analysis for 2017 was executed in 0.34 seconds, and the analysis for 2018 in 0.29 seconds. 
This increases the code's efficiency by a factor of 2.5

![refactor_2017]()

![refactor_2018]()

## Results

The results show that DQ traded 35,796,200 shares in 2017 and 107,873,900 in 2018. Considering those numbers, the value
of DQ stock has increased, and it has become more competitive. It is also the stock with the greatest return in 2017. However, 
when we consider DQ returns in both years, it went from 199.4% to -62.6%. This is a significant drop in the stock's value. 
This is why Steve's parents should consider diversifying their funds to mitigate the risks taken when investing. Also, 
as we can notice, DQ was not the only stock whose values dropped in 2018. This might suggest a bad year for green energy stocks,
while 2017 was perhaps a good year. Two stocks manage to get a positive return in 2018 ENPH and RUN. So, Steve's parents 
should also consider them as an investment option.

As for the efficiency of the code, when we were using one variable to record the total volume, the starting and ending price of each
stock, the analysis for 2017 and 2018 got executed in about 0.99 seconds. when we refactor the code to use arrays instead, 
the  analysis for 2017 and 2018 got executed in about 0.35 seconds.

## Summary
We refactor a code to improve it while keeping the same features or functionalities. So, the goal is to 
restructure the code to make it more reliable and remove bugs. Some of the advantages of code refactoring are: 
improving its design, maintaining it, and making it run faster. However, as refactoring a code involves modifying it, 
this can be risky and introduce other bugs. It can also be time-consuming.
