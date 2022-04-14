# Stock Analysis
Performing stock analysis using refactored VBA macro 
## Overview of Project
The purpose of this project is to analyze the stock market for years 2017 and 2018.  The excel file includes the stock ticker, open and close daily prices as well as the total number of daily traded shares for a number of stocks. A VBA macro is developed to calculate the yearly volume and the return percentage for each stock. The developed macro is a refactored version of the already-developed VBA script and therefore is expected to do the calculations faster than the original one.
# Analysis Results
## Stock Data Analysis
The VBA macro reads in the year for which the user wants to analyze the stock data for and then calculates and tabulates the total volume and return for each stock for year 2017 and 2018 as shown in Figures 1. The VBA macro also formats the calculated data as shown in the Figure. Return values less than 0% (loss) are colored in red while the returns greater than 0% (gain) are colored in green.  In general, the stock market performed much better in 2017 with the average gain of about 70% compared to average loss of about 9% in 2018. TERP was the only stock that lost its value in both years while ENPH and RUN stocks reported an average yearly gain of 105% and 45%, respectively. 

<div align="center"> 

![image](https://user-images.githubusercontent.com/103223944/163464933-f4bc81fa-89f8-441f-bf63-66f361e0cba8.png)
  
**Figure 1: Stocks total volumes and returns for 2017 and 2018** 
  
<div align="left"> 
  
## Refactored Macro Performance 
The original VBA macro developed through the VBA course is refactored for better computational performance. The major difference between the original and refactored macros is the way that the data for each stock is retrieved from the excel sheets. While the original macro loops over the entire set of data 12 times to calculate the yearly volume and return for all the stocks, the refactored macro loops over the entire set of data only once. This reduces the computational time by a factor of 4-5 as shown in Figures 2 and 3 below. 
  
<div align="center"> 
    
![image](https://user-images.githubusercontent.com/103223944/163465107-790d24c9-c112-40a4-8d4a-16ba99299f50.png)

**Figure 2: Runtime for refactored macro** 
      
![image](https://user-images.githubusercontent.com/103223944/163465188-4e58e087-d35e-425d-ab24-9845069457f7.png)

**Figure 3: Runtime for original macro** 
    
<div align="left">   
    
## Summary
In general, refactoring makes the code easier to understand, helps finding the bugs easier and also helps in executing the code faster. On the other hand, refactoring can potentially introduce bugs that cannot be easily caught specially when dealing with huge codes and therefore one needs to calculate the costs associated with the risk of introducing new bugs due to the refactoring.
For this project, the stock data was analyzed using a refactored VBA macro that reduced the runtime by a factor of 4-5. However, since the total runtime was only a fraction of second, the time spent on the code refactoring can only be justified by looking at it as a good practice!        
    
    
    
  
