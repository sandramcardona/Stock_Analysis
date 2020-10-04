# Stock_Analysis

## Overview of Project

### Objective of Project

To Analyze stock data provided by Client, Steve,  to create a summary of the stocks, total daily volume and yearly return for each stock and to format it so the client can easily spot stock performance at a glance and to be able to clear the data and re-run the report for different years. 

### Project Background

Steve's Parents have asked him to analyze the stock DQ, DAQO, and to provide how actively DQ was traded in 2018. Steve's parents belive that if a stock is traded often, then the price will accurately reflect the value of the stock.  Steve was asked to provide the total daily volume and yearly return for each stock on 2018 in the Green Stock data set. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.

Independent of the findings for DQ, Steve would also like to analyze multiple stocks to find other choices for his parents to invest in.
Steve may also want to look at a different set of stocks in the future so the program will be flexible for running multiple stocks and other years. 
Steve would like to be able to read the analysis at a glance so additional formatting will be added so the data is well organized and color coded to differantiate between positive and negative results at a glance. Buttons to run the macro and to clear the sheet will be added so Steve has access to run the code when needed without having to install or open additional windows. A timer for the program will also be added in case Steve needs to run it in larger sets, he can have an idea of how long it will take to run. 

## Results

The analysis is well described with screenshots and code (4 pt).
### Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

The analysis was performed on Green stock data for 12 stocks for 2017 and 2018. I started by creating 

![](/2017%20All%20Stoks%20Analysis_VBA_Challenge.png)



Now that we've run the analysis, let's make it easier for Steve to read by adding some formatting to our table. This is the same type of formatting we did in the last module—changing font styles, adding borders, setting number formats, and so on—but we can automate formatting with VBA.
Let's format our data so that Steve can determine stock performance at a glance.
Now that we've written and tested a significant amount of code, running a macro might seem like a simple task. But this might not be the case for Steve. He wants to focus on financial analysis, not installing Developer tools, determining the correct macro to use, and then figuring out how to run the macro. To make life easier for the end-users of our code like Steve, we can create buttons in the worksheet.
Steve will probably want to run this analysis for each year, so let's update our code to run for any year, not just 2018.
In the future, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results. To help Steve, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box.
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
