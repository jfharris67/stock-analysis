# stock-analysis
## Overview of Project
### Purpose

Steve is looking to expand the dataset of his stock market analysis to include more stocks. The current code works well for a small number of stocks, but it may not be efficient enough to handle a larger dataset. I will refactor the code to make it more efficient and improve its performance. I plan to test the refactored code to determine whether it is faster and more reliable than the original version

## Results

Step 1a: I declared a variable, tickerIndex, set to 0 – tickerIndex = 0

Step 1b: I created three output arrays:
 Dim tickerVolumes(12) As Long
 Dim tickerStartingPrice(12) As Single
 Dim tickerEndingPrice(12) As Single
 
Step 2a: I constructed a for loop to set the tickerVolumes to 0.

Step 2b: I built a for loop that iterates over all rows.

Step 3a: Within the for loop in Step 2b, I wrote a script that increases the tickerVolumes variable and adds the volume for the current stock ticker, utilizing the tickerIndex variable as the index.

Step 3b: I wrote an if-then statement that checks if the current row is the first row with the chosen tickerIndex. If so, then assigns the current starting price to the tickerStartingPrices .

Step 3c: I wrote an if-then statement that checks if the current row is the last row with the chosen tickerIndex. If so, then assigns the current closing price to the tickerEndingPrices.

Step 3d: I wrote a script that will increase the tickerIndex if the subsequent row's ticker does not match the preceding row's ticker.

Step 4: I utilized a for loop to iterate through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in the spreadsheet.

<img width="595" alt="Original 2017" src="https://user-images.githubusercontent.com/114355199/207204648-15f41a07-29d2-45b9-b4a1-857d41cc9e12.png">
<img width="599" alt="Original 2018" src="https://user-images.githubusercontent.com/114355199/207204773-0b9eb65c-7ad4-42e0-ac70-2437fdeb4fc6.png">
<img width="607" alt="Refactored 2017" src="https://user-images.githubusercontent.com/114355199/207204823-025307b2-6d83-4068-b7cf-973d04d3c538.png">
<img width="603" alt="Refactored 2018" src="https://user-images.githubusercontent.com/114355199/207204865-e2ec9f03-fc90-44c5-a5f4-208afc38f6d2.png">

##Summary

## Advantages and Disadvantages of refactoring VBA Scripts

### Advantages:

Refactoring VBA scripts can make the code easier to read and understand, which can make it easier to maintain and modify in the future.
Refactoring VBA scripts can improve the performance of the code, making it run more quickly and efficiently.
Refactoring VBA scripts can improve the reliability and stability of the code, reducing the risk of errors or bugs.

### Disadvantages:

Refactoring VBA scripts can be time-consuming and complex, requiring significant effort and expertise.
Refactoring VBA scripts can introduce errors or bugs, which can be difficult to detect and fix.
Refactoring VBA scripts can be challenging for developers who are not familiar with the VBA language or the specific features and capabilities of Excel.
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

## Pros and Cons Between Original and Refactored VBA Code
### Original Code
Advantages - It was a lot easier to understand with my limited knowledge. Took a lot less time to write. 
Disadvantages - The script had to run through all of the rows for each stock.

### Refactored Code
Advantages - The refactored version took a fraction of the time to run compared to the first version. This is because the original version looped through all the rows in the spreadsheet for each stock, while the refactored version collected all the information in a single loop. Although the difference may not be noticeable with a small dataset, it is important to optimize performance for larger datasets.
Disadvantages - The refactored script took a while to figure out. I am sure with more experience,my skill will improve.
