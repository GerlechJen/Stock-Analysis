# Stock Analysis with VBA
## Overview of Project 
A Finance degree holder, Steve, who wanted to assist his parents make the right decision regarding which  green energy stock to invest in used VBA for his analysis. His parents without much research, had decided to invest all their money in DAQO New Energy Corp stock. An Excel file that contained a handful of green energy stock data including that of DAQO New Energy Corp was analyzed by steve. He used the extension to excel, VBA, for his analysis so that the analysis could be automated and reused with any stock data. After his analysis, the result was that DAQO New Energy Corp had a return of 1.99 in 2017 but in 2018 its return dropped to -0.63.

### Purpose

In this project,the existing VBA code that was created by steve to help determine which green energy stock his parents should invest in will be refactored. Refactoring the code will allow the VBA Macro to run more efficiently. That is,it will be able to run thousands of stocks while using just a short time to execute. The goal here is to make it possible to expand the dataset to include the entire stock market over the last few years while running in a short amount of time.  


## Results

In making this analysis,

From the analysis, the overall stock performance of the year 2017 was better than the performance for 2018. The year 2017 had just one stock (TERP) having a negative return. While the year 2018 had just two stocks (ENPH and RUN) having a positive return 


![2017analysis](C:\Users\janno\Downloads\2017_analysis.png)

![2018analysis](C:\Users\janno\Downloads\2018_analysis.png)

When the original code steve created was run it took about 0.96 seconds to run the code for the year 2017 and about 0.89 seconds to run the code for the year 2018.

![2017image](https://github.com/GerlechJen/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2017.png)

![2018image](https://github.com/GerlechJen/stock-analysis/blob/main/RESOURCES/VBA_Challenge_2018.png)

## Summary 
Refactoring code has a lot of advantages. Some of the advantages include:
- It improves the logic of the code and makes the code easier to read and understand.
- It makes the code use less memory hence the is able to run faster .
- It makes the code reusable.
- It helps to improve the functionality of the code without adding any new functionality.
- It makes the code more reliable.

When it comes to disadvantages, if not carefully executed, refactoring could introduce new bugs or errors into the existing code.
How do these pros and cons apply to refactoring the original vba script 
If the pros and cons of refactoring are to be applied to the refactoring carried out on the VBA script for this project, it was realised that when the code had not been refactored, the run time for the 2017 analysis was 0.96 seconds and that of 2018 was  0.89 seconds. However, after refactoring, the code runtime reduced to about 0.19 seconds for 2017 and 0.16 seconds for 2018.
Also after refactoring the code if thousands of stock data for other years are to be included in the Excel file, the code will still be able to analyze that large data set efficiently. The initial code worked for the analysis of the two years and even if we had a dozen years' stock data to work with it would have still run but it might not run as well, if we try to analyze thousands of stock data. Even if it would run, it would take a very long time to execute the code.
Another thing that was observed was that for the initial starting code when a year is run, the formatting applied to colour the cells based on whether the return was positive or not does not work automatically on the next year analysed. For the other year analysed to have a correct colour formatting, the formatAllStocksAnalysisTable() macro has to be run again for the correct colours to be applied. When it comes to the refactored code the functionality was improved. When different years are checked using the refactored code the colour formatting works perfectly for any year analyzed without having to re-run the formatAllStocksAnalysisTable() macro.
