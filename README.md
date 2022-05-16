# Stock Analysis with VBA
## Overview of Project 
A Finance degree holder, Steve, who wanted to assist his parents make the right decision regarding which  green energy stock to invest in used VBA for his analysis. His parents without much research, had decided to invest all their money in DAQO New Energy Corp stock. An Excel file that contained a handful of green energy stock data including that of DAQO New Energy Corp was analyzed by steve. He used the extension to excel, VBA, for his analysis so that the analysis could be automated and reused with any stock data. After his analysis, the result was that DAQO New Energy Corp had a return of 1.99 in 2017 and a return of -0.63 in 2018.

### Purpose

In this project,the existing VBA code that was created by steve to help determine which green energy stock his parents should invest in will be refactored. Refactoring the code will allow the VBA Macro to run more efficiently. That is,it will be able to run thousands of stocks while using just a short time to execute. The goal here is to make it possible to expand the dataset to include the entire stock market over the last few years while running in a short amount of time.  


## Results

From the analysis, the stock performance of 2017 was better than the performance for 2018. 
When the original code steve created was run it took about 0.93 seconds to run the code for the year 2018 and about 0.98 seconds to run the code for the year 2017. 

## Summary 
Some advantages of refactoring code are that:
- It makes the code easier to read and understand.
- It helps to run the code faster .
- It makes the code reusable
- It helps to improve the functionality of the code and makes it more reliable.

When it comes to disadvantages:
- If not carefully executed, it could introduce new bugs or errors into the existing code.
How do these pros and cons apply to refactoring the original vba script 
Using this project for instance, when the code had not been refactored, the run time for the 2017 analysis was  and that of 2017 was  seconds. However, after refactoring, the code runtime reduced drasticallly to seconds for 2017 and for 2018.
Also refactoring the code implies if stock data for other years are to be included in out Excile file the code will still be ableto analyze that large dayta set. Initially the code was restricted for a specific year only .
