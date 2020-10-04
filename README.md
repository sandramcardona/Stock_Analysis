# Stock_Analysis

## Overview of Project

### Objective of Project

To Analyze stock data provided by Client, Steve,  to create a summary of the stocks, total daily volume and yearly return for each stock and to format it so the client can easily spot stock performance at a glance and to be able to clear the data and re-run the report for different years. 

### Project Background

Steve's Parents have asked him to analyze the stock DQ, DAQO, and to provide how actively DQ was traded in 2018. Steve's parents belive that if a stock is traded often, then the price will accurately reflect the value of the stock.  Steve was asked to provide the total daily volume and yearly return for each stock on 2018 in the Green Stock data set. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year which is how much the investment grew or shrunk by the end of the year.

Independent of the findings for DQ, Steve would also like to analyze multiple stocks to find other choices for his parents to invest in.
Steve may also want to look at a different set of stocks in the future so the program will be flexible for running multiple stocks and other years. 
Steve would like to be able to read the analysis at a glance so additional formatting will be added so the data is well organized and color coded to differantiate between positive and negative results at a glance. Buttons to run the macro and to clear the sheet will be added so Steve has access to run the code when needed without having to install or open additional windows. A timer for the program will also be added in case Steve needs to run it in larger sets, he can have an idea of how long it will take to run. 

## Results

The analysis is well described with screenshots and code (4 pt).

The analysis was performed on Green stock data for the DQ, DAQO, stock. The analysis was performed by looking for the DQ ticker inside the data. Once the DQ tickers were found there were two items to look for: the Daily Volume and the Yearly return.  
Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. For the analysis, all the volumes for the DQ tickers were added for the year 2018 to calculate the total volume and then it was displayed on the Total Daily Volume column. Next, for the Yearly Return, which is the percentage difference in price from the beginning of the year to the end of the year, the starting Price and the ending price for the DQ stock was substracted and then shown as a percentage in the Return column. 
Below are the results of the analysis for the DQ stock for 2018, showing that the performance of the stock did not do well so Steve needed to look for better stock options for his parents. 
![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/DQ_2018_Stock_Analysis.png)

Steve then wanted to see all the other stocks analysis to help his parents pick a better option. To run the analysis for all the stocks for 2017 and 2018, the code was edited to be able to be run for the specific year that Steve wanted to analyze. Instead of looking at one year a yearvalue variable was created so the year would be based on the input year Steve would add. Then an index was added to look for all the stocks instead of only DQ. Then different parameters were added so the code would recognize where the first line and last line of each stock was located in the list in order to add the correct Total Volumes and calculate the Yearly return correctly. This values were then placed into a table where the Stocks Year, Ticker name, Total Daily Volume and Return. The table was then formatted and the color was added to show in green any positive returns and red for any negative returns. 


Below are the images of the All stocks analyis for 2017 and 2018 based on the edited code. 
#### All Stocks Analysis 2017
![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/2017%20All%20Stoks%20Analysis_VBA_Challenge.png)

#### All Stocks Analysis 2018
![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/2018%20All%20Stocks%20Analysis_VBA_Challenge.png)

Based on the results for the all stocks analysis performed for 2017a nd 2018, the best stock to invest would be ENPH. In 2017, it had a percent return of 129.5% in 2018 it had an 81.9% return while the majority of the other stocks were in negative Returns. Eventhough the stock RUN had a low return in 2017 of 5.5%, it showed that in 2018 it had a high return while the other stocks had a negative return. RUN will also be a good option to invest in.  

Steve explained that he would like to use the code in other years for the same stock. In order to have the code make it run faster for when Steve runs larger data sets for these Stocks the code needed to be optimize.  In order to do this, the original code was refactored to use an Array instead of Ranges in order to improve the running time. This was performed by creating a tickerIndex variable and assigning 3 different output arrays. The tickerIndex was then added throughout the rest of the code to increase the total volume of the ticker and get the starting prices and ending prices for each ticker and then calculate the return from these two values. After completing each ticker the tickerIndex will then be increased by 1 so it would do the same for the next ticker. The rest of the formatting and color coding was kept the same. Below are the images from the original code and the running times for each year and then following that are the images for the refactored code and the new running times. 

##### Original code and running times
![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/Original_VBAcode_All_Stocks_Analysis.png)

![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/VBA_Challenge_originalcode_2017_runningtime.png)

![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/VBA_Challenge_originalcode_2018_runningtime.png)

##### Refactored code and new running times

![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/Refactored_VBA_code_All_Stocks_Analysis.png)


![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/VBA_Challenge_refactored_2017_runningtime.png)


![alt text](https://github.com/sandramcardona/Stock_Analysis/blob/master/Resources/VBA_Challenge_Refactored_2018_runningtime.png)

The different in times between the original and the refactored times shows that the optimization of the code really works and it works 6 to 7 times faster than the original code. 


### Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).


Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
To do this, nn input window was created and the code was change so it wouldn't look into 2018 but into the value for the year in the input window. as well as a run button and clear sheet button to and the results will be formatted to be easily analyzed. 

Now that we've run the analysis, let's make it easier for Steve to read by adding some formatting to our table. This is the same type of formatting we did in the last module—changing font styles, adding borders, setting number formats, and so on—but we can automate formatting with VBA.
Let's format our data so that Steve can determine stock performance at a glance.
Now that we've written and tested a significant amount of code, running a macro might seem like a simple task. But this might not be the case for Steve. He wants to focus on financial analysis, not installing Developer tools, determining the correct macro to use, and then figuring out how to run the macro. To make life easier for the end-users of our code like Steve, we can create buttons in the worksheet.
Steve will probably want to run this analysis for each year, so let's update our code to run for any year, not just 2018.
In the future, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results. To help Steve, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box.
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
