# Stock-Analysis

## Overview of Project
The purpose of this project is to refactor the code and measure the performance of the VBA script for stock analysis. The run time of the code is measured and compared with the original code to determine if code refactoring has reduced the code run time. 

## Results
As part of refactoring, the following changes were made to the code
 * Three output arrays were created for TickerVolume, TickerStartingPrice and TickerEndingPrice.
 * TickerIndex variable was created as a common iteration variable to access the values across the four different arrays: the tickers array and the output arrays.
 * Using the tickerIndex variable as the index, the current Ticker Volume was increased.
 * For loop was used to loop through the arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns. 
 Fig 1 shows how the output arrays and the tickerIndex are defined and used to increase the tickerVolume. 

![code screenshot](https://github.com/chinzjay/Kickstarter_Analysis/blob/main/code%20screenshot.png)
|:--:|
|Fig 1. Screenshot of the code|

![2017-NR](https://github.com/chinzjay/Stock-Analysis/blob/main/2017-NR.png)
|:--:|
|Fig 2. Run Time for the VBA Script before refactoring for the year 2017|

![VBA_Challenge_2017](https://github.com/chinzjay/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
|:--:|
|Fig 3. Run Time for the VBA Script after refactoring for the year 2017|

We can see from Fig 2.(code run time for year 2017 before refactoring) and Fig 3.(code run time for year 2017 after refactoring) that it took 1.91 seconds for the code to run before whereas after code refactoring it took just 0.81 seconds. 

![2018-NR](https://github.com/chinzjay/Stock-Analysis/blob/main/2018-NR.png)
|:--:|
|Fig 4. Run Time for the VBA Script before refactoring for the year 2018|

![VBA_Challenge_2018](https://github.com/chinzjay/Stock-Analysis/blob/main/VBA_Challenge_2018.png)
|:--:|
|Fig 5. Run Time for the VBA Script after code refactoring for the year 2018|

We can see from Fig 4.(code run time for year 2018 before refactoring) and Fig 5.(code run time for year 2018 after refactoring) that it took 1.93 seconds for the code to run before whereas after refactoring it ran in just 0.86 seconds. The run time in both the cases has come down by more than half. Therefore we can conclude that code refactoring makes the code run more efficiently. 

## Summary
Refactoring the code has advantages but at the same time it can also have some disadvantages.
Some of the advantages of refactoring the code includes
 * Makes the code efficient
 * Improves the logic
 * Makes the code more readable
 * Decreases the code run time
 
Some of the disadvantages of code refactoring are
 * It is time consuming to fix the errors if any are created during the process
 * If not done correctly, it may break the code

In our case, the creation of the iteration variable tickerIndex made the code more readable. It was used to loop through one time and collect the data. We can see from the results that the run time has reduced significantly after code refactoring (almost half!). It took some time to refactor the code and to fix the bugs that were created during the process. omparing to the pros, we can say that code refactoring plays a key role in making the code more efficient.

