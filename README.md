# stocks-analysis

# Investment Stocks Analysis

## Overview of Project

### Purpose
Steve has requested an Excel workbook that will allow him to analyze past stock market performance in order to determine the best stocks for Steve's parents to invest.
The Excel workbook will provide a simple interface for Steve to easily input the year he would like to analyse and output the Total Daily Volume and return for each stock in the dataset.


## Results

### Stock Performance Results
Looking at the results of the analysis for the [2017](/Resources/VBA_Challenge_Stocks_2017.PNG) and [2018](/Resources/VBA_Challenge_Stocks_2018.PNG) year indicates a decline in the stocks in 2018. The 2017 year showed positive investment growth for the majority of stock prices but the 2018 year showed a decline in most of the stocks.

The "ENPH" and "RUN" stocks performed well over both the 2017 and 2018 years showing a positive return for both years.
the DQ stock which Steve's parents invested in showed a very good 199.4% return in the 2017 year but unfortunately the 2018
year resulted in a 62.6% loss. This loss did not result in Steve's parents losing their initial investment but has caused a delay
in the long term investment growth.

![2017](/Resources/VBA_Challenge_Stocks_2017.PNG) ![2018](/Resources/VBA_Challenge_Stocks_2018.PNG)

### Execution Time Comparison
Looking at the execution time of the analysis between analysing the original code compared to the refactored code shows about an 85% reduction in processing time for both years.
The processing time for this dataset is still below one second but the more efficient refactored code would make a large difference in larger datasets.

![2017 Original Code Runtime](/Resources/VBA_Challenge_2017_Orig.PNG) ![2017 Refactored Code Runtime](/Resources/VBA_Challenge_2017.PNG) 

![2018 Original Code Runtime](/Resources/VBA_Challenge_2018_Orig.PNG) ![2018 Refactored Code Runtime](/Resources/VBA_Challenge_2018.PNG)

The Reason for the improvement after refactoring the code is a result of changing the for loop to only go through the dataset once instead of going through the dataset for each stock ticker.
Comparing the [original for loop](/Resources/VBA_Challenge_Original_forloop.PNG) to the [refactored for loop](/Resources/VBA_Challenge_Refactored_forloop.PNG) it can be noted not only is the code shorter but there is no longer a nested for loop, instead conditional if statements are used on each row of the dataset in the refactored code to analyse the data for each row of the dataset.

The original code shows a nested for loop looping through the entire dataset for all 11 stock tickers. 
 
The original code used a nested for loop looked like the following:  
note: the below is a high level code description, to view the whole code sample with code comments refer to the following - [original for loop](/Resources/VBA_Challenge_Original_forloop.PNG)
```
For i = 0 To 11
	totalVolume = 0
	ticker = tickers(i)
		Worksheets(yearValue).Activate
	For j = 2 To RowCount
		If Cells(j, 1).Value = ticker Then
			totalVolume = totalVolume + Cells(j, 8).Value
		End If
		If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
			startingPrice = Cells(j, 6).Value
		End If
		If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
			endingPrice = Cells(j, 6).Value
		End If
	Next j
```

The refactored code only loops through the dataset once and runs an analysis of the data in each row removing the need to loop through the dataset again.  

The refactored code looks like the below:  
note: the below is a high level code description, to view the whole code sample with code comments refer to the following - [refactored for loop](/Resources/VBA_Challenge_Refactored_forloop.PNG)
```
For i = 2 To RowCount
	If Cells(i, 1).Value = tickers(tickerIndex) Then
		tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
	End If
	If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
		tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
	End If
	If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
		tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
		tickerIndex = tickerIndex + 1
	End If
Next i
```


## Summary
When working with code sometimes the logic used throughout the initial development can result in inefficient code.
Refactoring code allows a developer to look at the code once it is complete and find ways to make it more efficient.

### Advantages and Disadvantages of Refactoring Code:

Advantages:
 - Allows developers to go back and add notes and comments to better explain what is happening in the code
 - Allows developers to make the code more efficient
 - Allows Developers to move code around to make it easier to find and fix bugs

Disadvantages:
 - Refactoring code takes time to complete and may extend timelines to complete a project
 - Due to the extended time to complete a project may result in additional costs

### Comparing Refactored Code to Original:
Refactoring the original code used in the analysis allowed the analysis to be more efficient and consume less resources.  
The original code had the following disadvantages:
 - The code looped through the data set multiple times consuming more resources than necessary
 - The code took longer to run than necessary

The refactored code showed the following advantages:
 - The amount of time to run the analysis was quicker
 - Consumed resources were less as the analysis only looped through the dataset once
 - the code was shorted and neater making it easier to understand
