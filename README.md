# stock-analysis
Katelynn Youssefyeh
Module 2 Challenge: VBA of Wall Street
Overview of Project
	The purpose of this analysis was to use VBA to analyze stock market data to determine which green energy stocks are worth investing in. Using VBA script, we can calculate the “Total Daily Volume” and “Return” of each stock without having to manually calculate it within the Excel spreadsheet. Specifically, we refactor the solution code to make the process more efficient. 

Results
	Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

After running the refactored script, we can see a clear difference between how the stocks performed in 2017 and 2018. To determine this, we primarily look at the return of each stock. First, we had Excel determine the starting and ending price of each ticker so that we can use those factors to calculate the return. For example, to find the ending price I used the code “If Cells(i + 1, 1).Value <> tickers(tickerindex) Then tickerEndingPrices(tickerindex) = Cells(i,6).Value” Once the starting and ending prices have been determined, we can calculate the return and put it into the “Returns” column of the spredsheet by using the code “Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) – 1” To make it easier to determine which stocks performed well, we also used VBA to format the excel sheet so that the positive returns would be highlighted in green, and the negative ones in red. Below are the results for 2017 and 2018: 


As we can see, the stocks in question did significantly better in 2017 compared to 2018.
	Refactoring is a process used in order to make your code more efficient and therefore fun faster. Thanks to the timer coded in, after refactoring the script we can see how long each analysis took for each year before and after refactoring the script. Below are the times for how long it took to run the original script:




After refactoring the script, the analysis became faster by about 0.7 seconds for both years.

















Summary
	When refactoring code, it’s important to be aware of the advantages and disadvantages. One advantage that we’ve seen is that it can make running your code more efficient and run faster. However, a disadvantage is that it can lead to more bugs or mistakes in your script. For example, when trying to refactor the code, I missed typing “(i)” when assigning the values to the cells. As a result, my code would not run until I realized my mistake. 
