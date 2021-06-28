# Stock Analysis using VBA
## Overview of Project
The purpose of this project is to determine the total daily volume and yearly return for a set of given stocks. 
Daily volume definition: Total number of shares traded throughout the day and provides idea on how actively a stock is traded. 
Yearly return definition: Percentage difference in price from the beginning of the year to the end of the year.
## Results
It appears that in 2017, from yearly return perspective DQ stock performed best and could be the best choice for people to invest.
However, in 2018, most of the selected stocks performed worst from return perspective other than ENPH and RUN

Only stock TERP gave negative return in 2017 although the daily volume is higher than DQ. 

Following snapshot shows the comparison of listed stocks between 2017 and 2018:
![StockComparison_2017_vs_2018](https://user-images.githubusercontent.com/62515666/123532228-16905a00-d6d1-11eb-84f9-16c53c33d080.png)
Following table shows the run time comparison of original vs. refactored script:
![RunTime_Comparison](https://user-images.githubusercontent.com/62515666/123532237-2c058400-d6d1-11eb-8060-4e86efe3a10e.png)
Below are the runtime snapshots of the refactored script:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/62515666/123532247-3fb0ea80-d6d1-11eb-9c37-dc6496832164.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/62515666/123532257-4d667000-d6d1-11eb-83bc-87c5c4aaa641.png)
## Summary
1. Refactoring plays an important role for computer programmer. It makes software easier to understand, helps finding bugs and also helps in executing the program faster. 
However, refactoring has certain disadvantages as side effects, for example, developer and tester needs to do more testing and subsequently it becomes more time consuming.
2. In our original VBS script (without refactor), the excel sheet was read multiple times for each stock whereas in the refactored script, the sheet was read only once to scan through all the stocks. This makes the execution time significantly less. However, it takes longer time for anyone to write the refactored script.
