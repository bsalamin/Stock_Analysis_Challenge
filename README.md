# Refactoring Stock Analysis in VBA & Excel

## Objective
The purpose of this project was to refactor the VBA Module in the green_stocks.xlsm so that way Stave can use it to analyze an expanded dataset to include the entire stock market over the last few years. Ideally, the refactored macro should run faster than first version, which took about 1 second to run the analysis.

## Process Overview
The refactored VBA module encompassed the following steps:
1.	The tickerIndex was set to 0, before looping over all the rows.
2.	3 arrays were created for the stock tickers, volumes, and starting and ending prices.
3.	The tickerIndex accesses the stock ticker index for the 3 arrays.
4.	The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
5.	The script organizes and formats the data into a table in the ‘All Stocks Analysis’ tab in the worksheet.
6.	Finally, the macro is assigned to the ‘Run Analysis’ button in the ‘All Stocks Analysis’ tab, which prompts the user to enter the year that they would like to analyze stocks for, and displays the script’s runtime in a message box.
<img width="521" alt="Code_1" src="https://user-images.githubusercontent.com/100387078/158082792-05fa5a41-e26e-4896-af78-81b6392db743.png">
<img width="620" alt="Code_2" src="https://user-images.githubusercontent.com/100387078/158082793-6bb65c14-29e8-43cd-b4ed-6718f2d3ff14.png">
<img width="603" alt="Code_3" src="https://user-images.githubusercontent.com/100387078/158082794-683244c9-4bb8-4d65-bc36-384da31d8374.png">

## Results
The script correctly analyzes either stock index, based on the user input, and formats the table correctly. The runtime to loop through either year’s stock ticker is now at or below 0.65 seconds, which is a 35% reduction over the script we started with.
![2017 Stock Analysis_Results](https://user-images.githubusercontent.com/100387078/158082855-cd7cfc78-087f-4eec-af11-c5184fdcdfb9.png)
![2018 Stock Analysis_Results](https://user-images.githubusercontent.com/100387078/158082858-e9ebe38e-dea3-446b-8715-f8e39e6912e1.png)

## Summary
### Challenges of Refactoring Code in General
* Refactoring code must be performed in small steps, or you risk breaking a script that works. The adage “if it ain’t broke, don’t fix it” may be considered when choosing whether or not to refactor code.
* Refactoring can improve code’s maintainability. If the code is easier to read and it’s intended purpose is clearer, it will be easier to maintain the code and debug it when necessary.
* You can extend the capabilities of the code more easily. If you refactor code to have a more patterned, clarified design, you can more easily extend the capabilities of the code and provide greater flexibility in its function.
*But, refactoring deteriorates the existing structure or architecture of the source code. Attempting a refactor in order to improve the codes maintainability or extensibility may have the opposite affect and/or reduce the comprehensibility of the code’s architecture or purpose.

## Challenges of Refactoring This VBA Script
### Refactoring this script: 
* Improved the comprehensibility of the code’s design.
* Extended the capabilities of the main script so as to reduce the number of overall macros that a user needed to run in order to achieve the same results – e.g., cell formatting.
* Did not expand the code’s ability to take in new or different stock tickers. Future analysis outside of the existing stock tickers in the dataset will require refactoring the code again.
