# Stock-Analysis

# An Analysis of Stock Data from 2017 and 2018

## Performing Stock Analysis To Understand How Certain Stocks Performed In Previous Years

Purpose of the analysis is to find out if these stocks are worth investing in by reviewing their total volume and yearly return for both 2017 and 2018. By understanding if a specific stock had a positive return with significant volume in 2917, and then into 2018 as well, we can possibly determine it to be a sound investment in the years to come. Two stocks that definitely stood out were ENPH and RUN. ENPH's return was 129.5% in 2017 and 81.9% in 2018. Run's return was 5.5% in 2017 and 84.0% in 2018 while many others had a negative return like DQ, HASI, and SPWR.

## Analysis Refactoring

The refactored VBA script allowed our analysis to run at a slightly quicker rate. The execution time for the original script was 0.640625 in 2017 and 0.6523438 in 2018. After refactoring the code, the execution time improved with a run time of 0.625 in 2017 and 0.6210938 in 2018.

### Advantages of refactoring code

The advantages of refactoring code, especially, in VBA is that it can efficiently run with faster execution times. This would be especially important if a program used a significantly large amount of code - the quicker run times for each sub or script the more efficiently the program may run. Refactored code may also be cleaner to read and easier to edit or apply to another data set.

### Disadvantages of refactoring code

Some disadvantages may be that refactoring code incorrectly can change its external behavior. By attempting to refactor existing/working code you may run into issues that require frequent debugging. If you're in a scenario at work with a specific deadline, this may limit the amount of time to actually test the code or product.

### How these pros and cons apply to refactoring the original VBA script for our stock analysis

We can evidently see the run time decreased from our original code and refactored code. The script decreased the run time by 0.015625 for 2017 and 0.03125 for 2018. If this was presented to a user, the user would most likely prefer a program that runs quicker while retreiving the same information. Although we succesffully refactored the code, it took time to troubleshoot and debug issues such as expected array and overflow errors. Given this, through trial and error I was able to better understand how the code worked and could see how it can be manipulated to a larger data set. 
