# Stocks Analysis

## Overview of Project
Automating stock analysis of several stocks using VBA. The script was developed through progressive refactoring and measuring performance.

## Results
In the original macro, the way it was looped forced the computer to go back and forth between locating a value for the volume of a particular stock, and adding it to the total for that stock's volume on the All Stocks Analysis worksheet. The back and forth nature of the loop caused it to perform slowly. The below images are the times for how long it took the macro to perform the analysis on each of the two years being pulled from.

![Original_All_Stocks_Run_2017](https://user-images.githubusercontent.com/102050273/194398083-ff854f99-f538-40f4-a3fe-b8e416eff5f1.png)

![Original_All_Stocks_Run_2018](https://user-images.githubusercontent.com/102050273/194398118-de12c34f-b1c4-44fc-a963-05c64d6e93a0.png)

To imporve the times, which in turn will allow the macro to handle larger quantities of data, it needed to be refactored. This was done by implementing an index *tickerIndex = 0* for the ticker names given to each stock, and creating an array for the three necessary outputs using *Dim tickerVolumes(12) As Long*, *Dim tickerStartPrices(12) As Single*, *Dim tickerEndPrices(12) As Single* . Using the index and the array the data that we need for the All Stocks Analysis worksheet and can be looped through more efficiently. The 'for' loop used is very similar to the one in the original macro, except instead of looping through the original source of data, it loops through the arrays. Below are the images of the times for how long it took the macro for each year after the changes had been made.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/102050273/194398280-beb15fe9-0385-46c6-b512-8f51994dcdc5.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/102050273/194398303-7bae4c44-3f9d-4c12-af1a-f452a0e1972e.png)

## Summary
1. What are the advantages or disadvantages of refactoring code?
	- pros: makes code faster. Preserves a clean and maintainable architecture in evolving code. Reduces bugs.
	- cons: No additional functionality. Costs development time. Quality dependent on previous developers work.

2. How do these pros and cons apply to refactoring the original VBA script?
	- The execution time has improved and the code is clearer and easier to adjust for future updates.
	- The refactored script does not give ability to analyze any set or the whole set of existing stocks. The functionality is the same as the original VBA script which is to analyse the performances of the set of 12 green stocks.
	- Recoding the refactored script would be needed to give our client the ability to analyze any set of stocks.
