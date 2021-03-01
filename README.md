# stock-analysis
02 - VBA Module - Berkeley 2021

## Overview of the project
Creating a macro in VBA to help a client analyze an entire dataset of stocks for the years 2017 and 2018. The dataset has 3,013 entries per year.
Analysis needed:
 - Total volume per year of a specific stock
 - Return of a stock per year (comparing closing price on the first day of the year against the price on the last day of the year)

### Results
##### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
While 2017 saw a grow in the return of the stock for almost all companies (11 out of 12), 2018 cannot be considered as a successat all (only 2 out of 12).


<img width="215" alt="All Stocks 2017" src="https://user-images.githubusercontent.com/25446419/109451016-1083b880-7a12-11eb-86e8-7fd29d0d472d.png">
<img width="266" alt="All Stocks 2018" src="https://user-images.githubusercontent.com/25446419/109451017-111c4f00-7a12-11eb-8cb6-a3c2ee5148d7.png">


Also, I refactored my code achieving around 70% of efficiency when running the macro (lowered from .5 to .15 seconds). Results can be find in the Resources folder.

### Summary
#####  What are the advantages or disadvantages of refactoring code?
The advantages are that it runs faster using less power of your computer when using a large dataset.
One disadvantage may be that the code may end up messier and harder to understand. 
##### How do these pros and cons apply to refactoring the original VBA script?
One cons is that I had to create new arrays and delete a lot of the original script making it confusing at the beginning.
The advantage is that now it can run faster with a less powerful computer.
