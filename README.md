# Stock-Analysis
## Overview of Project
  The objective is to explore green energy stock performance by analyzing financial data using VBA. With the help of macros the data will be analyzed and with these data a conclusion is to be made as to whether or not a stock is worth investing or passing, just by seeing its behavior in a determined year. At last, a refactoring of the original code is implemented to compare the running time of both the original and the refactored code.

## Results
### All Stocks Analysis
  After implementing the code to run the analysis on 2017 green stocks it looks something like this:
  
  ![2017](https://user-images.githubusercontent.com/83614893/149453094-5eea34da-8193-4d39-aa56-83f124e895b9.png)
>Image 1. Stock Analysis for 2017  

```
  Overall all stocks performed neatly, with one exception (TERP). Although if a recommendation is to be made the 
  greatest numbers are from DQ, followed by SEDG, ENPH, and FSLR with over 100% of return.
```  
  
For 2018 came out as follows:

![2018](https://user-images.githubusercontent.com/83614893/149453106-4daa741a-34a8-4217-acd1-b1aaef2227a0.png)
>Image 2. Stock Analysis for 2018 

```
Comparing these numbers with the previous year, the only one with steady grow is ENPH with 129.5% in 2017 and 
81.9% in 2018 making it the most reliable stock of both years. RUN also had growth in both years but this 
year's could be descartable as is below 10%. All others took a deep dive into the negative numbers.
```


### Code Refactoring
After refactoring the original code the runtimes are as follows:

![runTime_2017](https://user-images.githubusercontent.com/83614893/149452742-69705604-caf2-4bdb-885a-8f910cebadd2.png)
![runTime_2018](https://user-images.githubusercontent.com/83614893/149452749-727071ec-979e-4068-8150-bac658c79c06.png)
>Images 3 and 4. Original Run Time (2017, 2018 respectively)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/83614893/149452770-78c56835-6ef5-4f5c-8820-1aeeccdb3038.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/83614893/149452777-235e0a6b-be38-44a3-94b3-5f82ec698781.png)
>Images 5 and 6. Refactored Run Time (2017, 2018 respectively)



## Summary
  1. What are the advantages or disadvantages of refactoring code?
  2. How do these pros and cons apply to refactoring the original VBA script?
