# Stock-Analysis

## Overview of Project
The purpose of this project is to refactor the code to measure the performance of the VBA script.

## Results
Three output arrays were created for Ticker Volume, Ticker Starting Price and Ticker Ending Price. A tickerIndex variable was created to access the values across the four different arrays: the tickers array and the output arrays. A For loop was created to initialize the ticker volume to zero. Using the tickerIndex variable as the index, the current Ticker Volume was increased and the ticker Volume was added. Using a if-then statement we check if the current row is the first row with the selected tickerIndex. If it is, then the current starting price is assigned to the tickerStartingPrices variable. Fig 1 is a screenshot on how the output arrays and the tickerIndex was created and how the tickerVolume is increased. 

![code screenshot](https://github.com/chinzjay/Kickstarter_Analysis/blob/main/code%20screenshot.png)
|:--:|
|Fig 1. Screenshot of part of the code|

Another if-then statement is used to check if the current row is the last row with the selected tickerIndex. If it is, the current closing price is assigned to the tickerEndingPrices variable. The tickerIndex is increased if the next row’s ticker doesn’t match the previous row’s ticker. For loop is used to loop through the arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns in ythe Worksheet. 

![VBA_Challenge_2017](https://github.com/chinzjay/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
|:--:|
|Fig 2. Run Time for the VBA Script after code refactoring for the year 2017|

![2017-NR](https://github.com/chinzjay/Stock-Analysis/blob/main/2017-NR.png)
|:--:|
|Fig 3. Run Time for the VBA Script before refactoring for the year 2017|

Fig 3 represents the run time for the script before refactoring. Fig 4 represents the the run time after refactoring for the year 2018. It s evident from Fig 3 and Fig 4 that the run tim has been cut down to half after refactoring. 

![VBA_Challenge_2018](https://github.com/chinzjay/Stock-Analysis/blob/main/VBA_Challenge_2018.png)
|:--:|
|Fig 4. Run Time for the VBA Script after code refactoring for the year 2018|

![2018-NR](https://github.com/chinzjay/Stock-Analysis/blob/main/2018-NR.png)
|:--:|
|Fig 5. Run Time for the VBA Script before refactoring for the year 2018|

Similarly comparing Fig4 and Fig 5 we can conclude that the run time has reduced significantly (more than half) after code refactoring.

## Summary
Refactoring the code has its advantages as well as disadvantages.
Some of the advantages of refactoring the code includes
 * Makes the code efficient
 * Improves the logic
 * makes the code more readable
 * decreases the code run time
 
Some of the disadvantages of code refactoring are
 * It is time consuming to fix the errors if any are created during the process
 * If not done correctly, it may break the code

How do these pros and cons apply to refactoring the original VBA script?
It decreased the run time compared to the previous code. It took some time fixing the code.

