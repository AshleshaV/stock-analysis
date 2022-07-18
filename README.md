# Allstock_analysis
## Overview of Project  
 Steve requested us to help in running the analysis on the Green_stock data set for his parents to make a better choice in investing their funds. 
### Purpose 
The purpose of this project was to refactor our current stock analysis code.The initial request was to run a stock analysis code on all the stocks in the data set and to calculate the volume and the return value for all stocks. we build a macro and button for analyzing the data for all stocks in short time using Excel VBA code. 

Steve is happy with the work we prepared for the given data set, but it may take longer processing time to analyse thousands of stocks.so, Steve came with a change request to run the stock Analysis on thousands of stocks in a short time. the New request was to refactor the exisiting code to run the stock analysis on thousands of stocks with faster run time.

##  Results
When we ran the Initial analysis with the green_stocks .xlsm file the following run time was recordedas shown in the image.
For the 2017 all stock data.
![Greenstock 2017](![image](https://user-images.githubusercontent.com/96032051/179444308-7c671dc3-e033-424e-acb8-31d2035e599d.png)



For the 2018 all stock data.!
![Green stock 2018](![image](https://user-images.githubusercontent.com/96032051/179444342-f859db3d-8d49-431a-814c-ec14391d1c06.png)


In this challenge, after refactoring the code script, the code ran in ran in 0.0625 seconds for the 2017 stock data as shown in the image
![VBA_Challenge_2017](![image](https://user-images.githubusercontent.com/96032051/179444383-8f8634b6-8d3e-4cb2-9c7c-dd071c7e9129.png)
)




The code ran in ran in 0.0859375 seconds for the 2018 stock dataas shown in the image
![VBA_Challenge_2018](![image](https://user-images.githubusercontent.com/96032051/179444416-e571dce0-5da2-49fb-8e0d-170b5927d4a3.png)
)




#### Stock Performance comparision:
The data set of All stocks 2017 (VBA_Challenge_2017.png) shows all the tickers have a positive return range from 5.5%to 199.4%, except for the ticker TERP has a negative retun -7.2%.The data for the 2017 year shows the company perofrmed well than in 2018.
Instead of a nested loop which run all the variable by steps for a more number of times.In the refactored scrip We used a for loop to improve the run time, making separate loops. The original run time is 1 second and refactored script's run time is 0.0625 seconds.

The data set of All stocks 2018 (VBA_Challenge_2018.png) shows all the tickers have a negative return range from -3.5% to 62.6%, except for the ticker ENPH and RUN have return of 81.9% and 84.0%. So, for the year 2018 the data shows a negative returns comparitively to 2017.
refactored script's run time is 0.0859375 seconds while the original runtime is 1 second for the 2018 year stocks, As the refactored scrip shows a for loop to improve the run time, making separate loops.




## Summary

### What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code is code runs more efficiently with fatser run time for the larger data.
The disadvantage is that the code becomes more complex, time consuming to plan and work more on refactoring the current code.


### How do these pros and cons apply to refactoring the original VBA script?
The original script code is easy to write by creating a loop and put another loop under the first loop making it simple to write the code. However, the refactored script is more complex to write with different variables, which is time taking to plan and design the script code.


