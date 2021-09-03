# Stock-Analysis

## Overview of Project
The purpose of this project is to refactor the code to measure the performance of the VBA script.

## Results
Three output arrays were created for Ticker Volume, Ticker Starting Price and Ticker Ending Price. A ticker index variable was created to access the values across the four different arrays; the tickers array and the output arrays. A For loop was created to initialize the ticker volume to zero. Using the tickerIndex variable as the index, the current Ticker Volume is increased ans the ticker Volume is added. Fig 1 is a screenshot on how the output arrays and the tickerIndex was created and also how the tickerVolume is increased. 
https://github.com/chinzjay/Kickstarter_Analysis/commit/928e4af1b260ba26413bb53997d63029d6f5b4b3
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

Similarly comparing Fig4 and Fig 5 we can conclude that the run time has reduced significantly after code refactoring.
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

