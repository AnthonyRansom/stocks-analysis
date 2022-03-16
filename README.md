# stocks-analysis

# Investment Stocks Analysis

## Overview of Project

### Purpose
Steve has requested a Excel workbook that will allow him to analyze past stock market performance in order to determine the best stocks for Steve's parents to invest.
The Excel workbook will provide a simple interface for Steve to easily input the year he would like to analyse and output the Total Daily Volume and return for each stock
in the dataset.


## Results

### Stock Performance Results
Looking at the results of the analyis for the [2017](/Resources/VBA_Challenge_Stocks_2017.PNG) and [2018](/Resources/VBA_Challenge_Stocks_2018.PNG) year indicates a decline in the stocks in 2018.
the 2017 year showed a good majority in stock prices but the 2018 yeat showed a decline in most of the stocks.

The "ENPH" and "RUN" stocks performed well over both the 2017 and 2018 years showing a positive return for both years.
the DQ stock whcih Steve's parents invested in showed a very good 199.4% return in the 2017 year but unfortunately the 2018
year resulted in a 62.6% loss. This loss did not result in Steve's parents losing thier initial investment but has caused a delay
in the long term investment growth.

![2017](/Resources/VBA_Challenge_Stocks_2017.PNG) ![2018](/Resources/VBA_Challenge_Stocks_2018.PNG)

### Execution Time Comparison
Looking at the executiion time of the analysis between analysiing the 2017 year compared to the 2018 year doesn't show any major differenct between the analysis processing time
the 2017 year analayis took 0.125 seconds and the 2018 year took about 0.117 both analysis taking less than 1 second to run.

![2017](/Resources/VBA_Challenge_2017.PNG) ![2018](/Resources/VBA_Challenge_2018.PNG)

## Summary
When working with code sometimes the logic used throughout the initial development can result in ineffecient code.
Refactoring code allows a developer to look at the code once it is complete and find ways to make it more effecient.

### Advantages and Disadvantages of Refactoring Code

Advantages:
 - Allows developers to go back and add notes and comments to better explain what is happening in the code
 - Allows developers to make the code more effecient
 - Allows Developers to move code around to make it easier to find and fix bugs

Disadvantages:
 - Refactoring code takes time to complete and may extend timelines to complete a project
 - Due to the extended time to complete a project may result in additonal costs

### Comparing Refactored Code to Original:
Refactoring the original code used in the analysis allowed the analyis to be more effecient and consume less resources.
The original code required the process to go through the entire dataset for all 11 stock tickers compared to the refactored code which only went through the dataset once
and then went through each stock ticker for each row. This is more effecient on memory as less data needs to be reloaded into each memory for each loop in the for loop.
