# Stock Analysis with VBA
## Overview of Project 
A Finance degree holder, Steve, who wanted to assist his parents make the right decision regarding which  green energy stock to invest in used VBA for his analysis. His parents without much research, had decided to invest all their money in DAQO New Energy Corp stock. An Excel file that contained a handful of green energy stock data including that of DAQO New Energy Corp was analyzed by steve. He used the extension to excel, VBA, for his analysis so that the analysis could be automated and reused with any stock data. After his analysis, the result was that DAQO New Energy Corp had a return of 1.99 in 2017 but in 2018 its return dropped to -0.63.

### Purpose

In this project,the existing VBA code that was created by steve to help determine which green energy stock his parents should invest in will be refactored. Refactoring the code will allow the VBA Macro to run more efficiently. That is,it will be able to run thousands of stocks while using just a short time to execute. The goal here is to make it possible to expand the dataset to include the entire stock market over the last few years while running in a short amount of time.  


## Results

From the analysis, the stock performance of 2017 was better than the performance for 2018. 
When the original code steve created was run it took about 0.96 seconds to run the code for the year 2017 and about 0.89 seconds to run the code for the year 2018.

## Summary 
Some advantages of refactoring code are that:
- It improves the logic of the code and makes the code easier to read and understand.
- It makes the code use less memory hence the code runs faster .
- It makes the code reusable
- It helps to improve the functionality without adding any new functionality.
- It makes the code more reliable.

When it comes to disadvantages:
- If not carefully executed, it could introduce new bugs or errors into the existing code.
How do these pros and cons apply to refactoring the original vba script 

Using this project for instance, when the code had not been refactored, the run time for the 2017 analysis was 0.96 seconds and that of 2018 was  0.89 seconds. However, after refactoring, the code runtime reduced to about 0.19 seconds for 2017 and 0.16 seconds  for 2018.
Also refactoring the code implies if thousands of stock data for other years are to be included in our Excel file, the code will still be able to analyze that large data set. The initial code worked for the analysis of the two years and even if we had a dozen years' stock data to work with  it would have still run but it might not run as well, if we try to analyze thousands of stock data. Even if it would run, 
Another thing that was observed was that for the initial starting code when a year is run, the formatting applied to colour the cells based on whether the return was positive on on year does not work automatically on the next year analysed. for one year does not work wit would take a very long time to execute the code.  the formatAllStocksAnalysisTable() macro has to be run again for the correct colours to be applied. While when different years are checked using the refactored code everything including values and colour formatting works perfectly for any year amnallyzed without rerunning the Sub formatAllStocksAnalysisTable() macro.
