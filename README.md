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
1. Year 2017 - 3,012 records
2. Year 2018 - 3,012 records

The first version of our code was iterating the whole data set, once per ticker (stock symbol). That is equivalent to have a data set of 36,144 records (3,012 records x 12 tickers). Although the first iteration of our code was doing its job, it was not very efficient. There should be a way to de a better job.

### Results for first version of our code
Running our code for the 12 stocks, year 2017, took 0.5234 seconds to complete the analysis:
![Stock analysis for 2017 - first version (unrefactored)](/resources/stock_analysis_2017_unrefactored.png)

The same unrefactored version, for year 2018, took 0.7382 seconds to complete the analysis
![Stock analysis for 2018 - first version (unrefactored)](/resources/stock_analysis_2018_unrefactored.png)

### Results for the second version (refactored) version of our code
What we did to improve the efficiency of our code, was the following:
1. We created a set of data structures to hold the results of our analysis: an array to hold the tickers, a second array to hold the volumes, a third one to hold the initial prices, and a third one to store the closing prices. 
2. We ran through the data set only once - we were able to detect a change in the ticker by analyzing the data in the next and previous rows. This way, we knew if it was necessary to change our point to a different stock ticker. 
3. By the end of the execution, the needed information was stored in the four arrays we created. it was only a matter of printing out the results.

By doing these tweaks, we were able to dramatically increase the efficiency of our code. Now, instead of iterating the data set 12 times, only one iteration was enough. 

The execution time for the 2017 analysis was reduced to 0.125 seconds. That is a 24% time reduction. This is good.
![Stock analysis for 2017 - second version (refactored)](/resources/VBA_Challenge_2017.png)

The execution time for the 2018 analysis was reduced to 0.1445 seconds. That is a 19% time reduction. This is also a good improvement. 
![Stock analysis for 2017 - second version (refactored)](/resources/VBA_Challenge_2018.png)

## Summary and conclusion
### What are the advantages or disadvantages of refactoring code?
It really depends on the experience of each individual. It also has to do with every person's programming style. From my experience, I can tell that:
1. **Advantage:** Most of times, the problems we face when coding have been solved by someone else. It is just a matter of getting that code and adapt it to our own needs (re-usability). 
2. **Advantage:** We can save time by juts adapting somebody else's code.
3. **Disadvantage:** If the code is not well documented, and it is hard to understand, it will be hard to re use that particular piece of code. In this cases, most developers will choose to start from scratch. 
4. **Disadvantage:** If we find code and it works without any modification, we may not take the time to fully understand and document the code. Over time, this will not help the development team. 

### How do these pros and cons apply to refactoring the original VBA script?
- Advantage 1 (re-usability) - We were able to utilize a piece of code that was partially complete. We just addedd a few lines to make it work. This helped us to save time and not desinging a new solution from scratch.
- Advantage 2 (Efficiency) - Adpating this code to our needs was quick.
- Disadvantage 1 (not well documented code) - This particular issue was not found. The code we needed to refactor was fully documented and easy to understand. But if otherwise, we would have to invest some time in reading the code and fully document it. 
- Disadvantage 2 (working code) - This issue was not present either. First, the code was well documented, so we knew from the beginnnig how it works. But, the code was not complete. We needed to complete the code in order to make it work.


