# Stock-Analysis

## Overview of Project

### Purpose 

The purpose of this analysis was to create a more efficient workbook capable of calculating the total volume and price change of all stocks within the stock market over a one year period.

## Results

By refactoring the Stock Analysis macro, the time that it takes for the workbook to analyze the 12 stocks in question was expedited. The refactored script was able to operate more efficiently as can be seen by the difference in time that it takes for the code to perform the analysis on each data set (2017 and 2018). Although the difference in code may seem too small to matterr, from ~0.77 to ~0.28 seconds in the 2017 analysis and from ~0.59 to ~0.07 in the 2018 analysis, the efficiency is better understood when seen as a percentage change, a reduction in time of roughly 63% for 2017 and 93% for 2018. This matters because although it is nearly impossible to distinguish the difference between these two codes without a timer, if the code were to analyze thousands of ticker symbols instead of only 12, the difference in time would be a much more noticeable.

  - Time it took code for 2017 data before refactoring:
  
![Time it took code for 2017 data before refactoring:](https://github.com/josem279/Stock-Analysis/blob/master/green_stocks_2017.png)

  - Time it takes code for 2017 data after refactoring:
  
![Time it takes code for 2017 data after refactoring:](https://github.com/josem279/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017.png)

  - Time it took code for 2018 data before refactoring:
  
![Time it took code for 2018 data before refactoring:](https://github.com/josem279/Stock-Analysis/blob/master/green_stocks_2018.png)

  - Time it takes code for 2018 data after refactoring:
  
![Time it takes code for 2018 data after refactoring:](https://github.com/josem279/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018.png)


## Summary

**1. What are the advantages or disadvantages of refactoring code?**

Some of the potential advantages of refactoring code include making code more efficient, adjusting code to more specifically meet the needs of certain analyses, and improving the code readability - by making it more simple or including comments that detail the code.

Some of the potential disadvantages that can occur with refactoring code on the other hand, include potentially affecting the functionality of the initial code and making the newly refactored script harder to read if it does not include comments or is poorly formatted.

**2. How do these pros and cons apply to refactoring the original VBA script?**

When comparing the Original code and this newly created refactored VBA script, we can see some that there were pros and cons to each way of conducting the analysis.

**Pros and Cons of the Original Code**

The original code was able to conduct an analyses of the stocks that were presented for the 2017 and 2018 years rather efficiently. It was written in a consise manner that involved one long loop that would output a formatted value of the total volume and change in starting and ending prices for each ticker in one go. Though it was effective at conducting this analysis, by having so many calculation nested within one another, the code ran longer that needed and to someone unfamiliar with the code may have been hard to read as it could be easy to lose track of which line of code belonged to each method. 

**Pros and Cons of the Refactored Code**

The refactored code is easier to read as it separates sections of code more cleanly. For instance, the tickerIndex variable and the loop used to read the tickerVolumes array are now separated from the For Counter loop used to run through all of the rows of data found in each tab. Additionally, the formatting is also cut out of the loop that runs through all rows. By splitting up these variables from the much larger For loop and If calculations, the macro run faster as it first stores the information in the corresponding arrays - ticker, volume, and starting/ending prices - **_then_** outputting it into corresponding cells rather than attempting to run all the If calculations for each line of data it may potentially output. Furthermore, the reduction in nested loops and calculations allows users better see the flow of the code and may prevent future mistakes and trouble reading the code - such as losing track of which line of code corresponds to which method or forgetting to close an If calculation.
