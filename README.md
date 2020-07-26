# Stock-Analysis

## Overview of Project

### Purpose

The purpose of this analysis was to create a more efficient workbook capable of calculating the total volume and price change of all stocks within the stock market over a one year period.

## Results

By refactoring the Stock Analysis macro, the time that it takes for the workbook to analyze the 12 stocks in question was expedited. 

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

When looking at this particular VBA script as an example of refactoring code, we can see that we were able to incorporate many of the potential pros previously mentioned:

- The refactored script is now able to operate more efficiently as can be seen by the difference in time that it takes for the code to perform the analysis on each data set (2017 and 2018). Although the difference in code may seem too small to matterr, from ~0.77 to ~0.28 seconds in the 2017 analysis and from ~0.59 to ~0.07 in the 2018 analysis, the efficiency is better understood when seen as a percentage change, a reduction in time of roughly 63% for 2017 and 93% for 2018. This matters because although it is nearly impossible to distinguish the difference between these two codes without a timer, if the code were to analyze thousands of ticker symbols instead of only 12, the difference in time would be a much more noticeable.

- The code is now easier to read as it does not contain as many nested if statements in our For Counter loops and separates sections more cleanly. 

