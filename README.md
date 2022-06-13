# VBA_Stock_Analysis_Challenge
Use VBA script to automate stock analysis

# Stock Analysis with VBA
## Overview of Project 

In this project, I used VBA to help Steve analyze green energy stocks so that he could give his parents stock investment advice. To do this, I need to loop through all the data and collect the total daily volume and annual return rate for each stock in 2017 and 2018 respectively. 

### Purpose
The purpose of this analysis is to use VBA script to find an efficient way to automate data extracing, calculating, and formatting, wchich is usually conducted in Excel manually. To make the code more efficient, I refactored the code to get all the data I need in only one loop.

## Results
### Original VBA Code
In the original VBA code, I used two loops to scan all the data in the file, returned the expected output in the "All Stocks Analysis" worksheet, and recorded the code running time with a timer. 
The code running time for 2017 is 1 second. The code running time for 2018 is 0.95 seconds. 

### Refactored VBA Code
In the refactored VBA code, I modified the original code, reduced the loops to one, and presented all the information in the "All Stocks Analysis" worksheet.  

The code running time for 2017 is 0.21 seconds. And the code running time for 2018 is 0.17 seconds. The refactored code is much faster than the original VBA code I wrote for the stock analysis. 

### Stock Performance in 2017 and 2018
The stock (DQ), picked by Steve's parents, had a good performance in 2017. However, in 2018, DQ's return rate tumbled from 199.4% to -62.6%. DQ would not be a good investment option for steve's parents. 

For all green energy stocks' performances in 2017 and 2018, the green energy stocks had a better performance in 2017. In 2018, only two stocks had positive returns: EHPN and RUN. ENPH outperformed most of the other green energy stocks in both years. Even though its return rate dropped from 129.5% to 81.9% in 2018, as the general performance of the green energy stocks in 2018 was worse than in 2017, EHPN will be a good choice for investment. In 2017, RUN had an average performance among the other green stocks. However, in 2018 RUN is leading the way in both trading volume and return rate. Both ENPH and RUN are good options for Steve's parents to invest in.   

## Summary 
### Advantages and Disadvantages of Refactoring Code
Refactoring code is the process of restructuring existing code while not changing its functionality.
>Refactoring is intended to improve the design, structure, and/or implementation of the software, while preserving its functionality. (Wikipedia)

The advantages of code refactoring are: 
- Improve the internal structure of the existing code
- Enhance code performance; reduce running time of the existing code
- Enhance readibility of the code

The disadvantages of code refactoring are:
- Affect the functionality of the original code
- Introduce new bugs into the code

### Apply Code Refactoring to the Original VBA Script
I encountered several errors as I had to refactor the code with different logic. But in the process of debugging, I got a better understanding of how VBA codes run. After I completed code refactoring, my code's performance enhanced a lot. The running time is five times faster than the original code now.
