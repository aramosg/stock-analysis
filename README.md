# Module 2 - Excel VBA - Stock Analysis with VBA

## Overview of Project
The purpose of this project is to analyze a set of data, representing the history for 12 stocks. The main purpose is to find what are the best stocks for the specified year by the user. 

We have been given a piece of VBA code that we need to refactor to perform the task described above. By refactoring this code, we will be able to:
1. Optimize the performance of our Excel file
2. Discover what are the best options (stocks) for our customers

By the end of this Project, we will learn:
1. The best stocks for investing
2. How to build more efficiente code
3. How to refactor code and its advantages/disadvantages 

## Analysis and results

### Analysis 
When we start analyzing this data set, we built a first version of our code to read throght the whole set of records. For our 12 stocks, we are talking about:
a. Year 2017 - 3,012 records
b. Year 2018 - 3,012 records

The first version of our code was iterating the whole data set, once per ticker (stock symbol). That is equivalent to have a data set of 36,144 records (3,012 records x 12 tickers). Although the irst iteration of our code was doing its job, it was not very efficient. There should be a way to de a better job.

### Results for first version of our code
Running our code for the 12 stocks, year 2017, took 0.5234 seconds to complete the analysis:
![Stock analysis for 2017 - first version (unrefactored)](/resources/stock_analysis_2017_unrefactored.png)

The same unrefactored version, for year 2018, took 0.7382 seconds to complete the analysis
![Stock analysis for 2018 - first version (unrefactored)](/resources/stock_analysis_2018_unrefactored.png)

## Summary and conclusion
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
