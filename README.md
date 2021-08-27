# stock-analysis

## Overview of Project
### The purpose and background 
Steve has the workbook I prepared but he wants it to be more efficient if there are thousands of stocks to analysis. The purpose of the project is to refactor the initial code, and make the VBA script run faster.

## Results
Stocks outperformed in 2017 but underperformed in 2018. The refactored script execution time decreases compared to original script. <img width="595" alt="Screen Shot 2021-08-27 at 3 09 00 PM" src="https://user-images.githubusercontent.com/88211298/131194137-6e817272-d82c-45b3-99e8-ef6f188001ca.png">
<img width="593" alt="Screen Shot 2021-08-27 at 3 09 12 PM" src="https://user-images.githubusercontent.com/88211298/131194142-feda83d8-28af-4f69-951a-1fef7fbcd5bc.png">
<img width="589" alt="Screen Shot 2021-08-27 at 3 16 56 PM" src="https://user-images.githubusercontent.com/88211298/131194144-f655d2a4-96c2-44e8-8a5b-bc9b24b0bbcf.png">
<img width="590" alt="Screen Shot 2021-08-27 at 3 17 07 PM" src="https://user-images.githubusercontent.com/88211298/131194145-10e0bc57-ec0b-41f0-88d9-7ee81e9669c2.png">

The time decreases due to the method I used for refactored is I only loop the whole sheet only once (for i= 2 to RowsCount), but the original script has nested loop for each tick and go over and over again to loop the whole sheet (for i=0 to 11| for j=2 to RowsCount).<img width="698" alt="Screen Shot 2021-08-27 at 3 14 41 PM" src="https://user-images.githubusercontent.com/88211298/131195583-d209cddd-1e87-4148-899e-ef9a932a7d75.png">

## Summary
### The advantages and disadvantages of refactoring code
The advantages of refactoring code are shorter time to execute, take fewer steps and easier for user to read. The disadvantages are refactoring is not always easy to do which means it may take much time to think how to refactor, search online for help. Try different ways to find the best solution also cost time.
### The advantages and disadvantages of the original and refactored VBA script
The refactored VBA script takes less time to get results, it's not that obvious in small data but it will show a big progress in large data, and of course it makes script easier to read. The disadvantages are it's designed for this particular script because we can see the 2017&2018 tickers sort from A to Z, it's easy to use tickerindex=tickerindex+1. But if tickers are random, we should take more steps like sort data first to use refactored script.
