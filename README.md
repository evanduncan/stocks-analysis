# Stocks Analysis

## Overview
The purpose of this project was to analyze stock data and to practice refactoring VBA code. The background for this project is a request for a workbook that can take stock market data for a number of years and output a summary of the performance of the individual stocks for each given year.

## Results
Nearly every stock did better in 2017 than it did in 2018.
2017:
<img width="250" alt="2017_stock_performance" src="https://user-images.githubusercontent.com/18372229/140688029-84787e9a-9199-494f-b311-c92a3dc14858.png">

2018:
<img width="250" alt="2018_stock_performance" src="https://user-images.githubusercontent.com/18372229/140688051-d978db49-19cf-4632-a236-14674489ecb5.png">

Additionally, the refactored code ran significantly faster than the original code.
Original:

<img width="425" alt="2017_stock_analysis_speed" src="https://user-images.githubusercontent.com/18372229/140688163-6c8d8445-0816-46f4-9a73-7c8d61f43540.png">

<img width="430" alt="2018_stock_analysis_speed" src="https://user-images.githubusercontent.com/18372229/140688179-bd700588-d48f-4ddd-b83c-c6e8250f90f3.png">

Refactored:

<img width="443" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/18372229/140688232-3a3bd307-cd73-45f1-b329-677359065cac.png">

<img width="429" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/18372229/140688234-360f7b0e-9067-420e-a68b-60cb4385ea57.png">

This difference in runtimes is most likely due to the nested for loop in the original code. This is resolved in the refactored code with separate for loops.
Original:

<img width="524" alt="Original_code" src="https://user-images.githubusercontent.com/18372229/140688821-f25d507b-70dc-4e7b-8dcb-d5b21b0b9466.png">

Refactored:

<img width="556" alt="Refactored_code" src="https://user-images.githubusercontent.com/18372229/140688838-ed617914-e26c-4f1e-bacb-4892f50a0ab8.png">

## Summary
From this we can see that there are advantages to refactored code in that, if done properly, it will likely run faster than someone's original code. There might also be disadvantages. For instance if someone was not dilligent in commenting their original code, it can foten be hard to improve upon. In this project the refactored code proved to be much faster than the original, meaning that as the data increases in size the refactoring will allow for more reasonable processing times. One drawback to the refactored code is that it is a little less intuitive than using a nested for loop to comb over a sheet of data.
