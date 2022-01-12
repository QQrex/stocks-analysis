# stocks-analysis
## **Purpose**

Help Steve refactor his code to analyze his stock data and see if the refactored code runs faster than the original code we wrote for him.

## **Overview of Project**

Using the code he provided:

1. Review the code that was present to us and edit if needed

3. Add code to find total volume, starting volume and ending volume of each ticker

4. Find % change of ticker volume each year

5. Compare code runtimes to see if refactored code runs faster

## **Analysis and Challenges**

**[Reviewing code]**

Steve has provided us with a started code to work with, but before we start adding new lines of code, we should review what was given to us.

![Start sub, timer and](https://user-images.githubusercontent.com/96326293/148659851-80318037-090d-47ef-8692-013a9b3c956e.PNG)
>Line 2 - Start subroutine called AllStocksAnalysisRefactored
>
>Line 3,4 - Declaring startTime and endTime variable as Single
>
>Line 6 - Creating an InputBox for yearValue of stocks we are looking at
>
>Line 8 - Start timer for subroutine

![Format output sheet](https://user-images.githubusercontent.com/96326293/148659982-de0a2661-98b1-4fcd-a5d7-96ef457cb5f6.PNG)
>Line 11 - Declaring which worksheet the headers will go
>
>Line 13 - For range "A1" we want value = All stocks + year inputted from Line 6
>
>Line 16 to 20 - Set headers for output data

![ticker array](https://user-images.githubusercontent.com/96326293/148660161-61f9e1a4-3ede-41ff-8e9e-b0078bb17e21.PNG)
>Line 23 - Declaring tickers array as string
>
>Line 25 to 36 - Set value of each ticker to array

![Formatting](https://user-images.githubusercontent.com/96326293/148660243-98bbdf28-914d-450f-b457-b5db056cefa4.PNG)
>Line 94 - Declaring active sheet as All Stocks Analysis
>
>Line 95,96 - Setting font style and borders for headers in output sheet
>
>Line 97,99 - Setting Number format and "autofit" for output data in output sheet
>
>Line 101,102 - Declaring dataRowStart = 4 and dataRowEnd = 15
>
>Line 104 to line 116 - Creating for loop to format color coding for output %
>
>>Line 104 - Loop through dataRowStart to dataRow End
>>
>>Line 106 to 108 - Green for values >0
>>
>>Line 110 to 112 - Red for everything else
>>
>>Line 114 to 116 - End if statment and color format loop
>
>Line 118 - End timer
>
>Line 119 - Set up message box to display for runtime of subroutine
>
>Line 121 - End subroutine

**[Adding New Code]**


Now that we reviewed the starting code, we can complete the rest of the subroutine to perform our analysis. First, we would want to declare any variables or arrays we would need.

![RowCount, dim for index, tickerVolumes, StartinPrice and EndingPrice](https://user-images.githubusercontent.com/96326293/148661289-5c607164-d34b-465f-b190-737af5010959.PNG)
>Line 39 - Set active sheet
>
>Line 42 - Set row count variable to last row in "A" column with data
>
>Line 45 - Creating ticker index variable and declare variable as byte
>
>Line 48 to 50 - Declaring output arrays of tickerVolume as double, tickerStartingPrices as single and tickerEndingPrcies as single


Next we would want to set up a for loop to run for each of our tickers starting at ticker 0.
We also need to set our tickerVolumes = 0 in order to add each ticker's volume.

![image](https://user-images.githubusercontent.com/96326293/148663022-5853acd6-5e6e-4ee1-9c28-ff3664bc2cc3.png)
>Line 54 - Start for loop at 0 and then loop to ticker 11 to finish.
>
>Line 55 - Set ticker = ticker(index) for easier reference
>
>Line 56 - Initialize tickerVolumes = 0


After looping for all tickers, we need to nest another loop for each ticker value to run through each row of our spreadsheet.

![Loop for all rows](https://user-images.githubusercontent.com/96326293/148663147-2d1129f9-2374-4475-8ae8-cc4491d0f2c9.PNG)
>Line 59 - Set active worksheet to yearValue again for loop to run through dataset
>
>Line 61 - Start for loop i to run loop through row 2 to RowCount


Now that our loops have been set up, we need to create if statements to add up total tickerVolumes. We also need to create if statement to find what the ticker start and end prices are.

![if statment for ticker volumes, start and end price](https://user-images.githubusercontent.com/96326293/148663517-c0aa0957-2113-40c3-8af4-c985841cfb7b.PNG)
>Line 64 to 68 - If statement for cell value (i,1) is equal to ticker then we would add cell value (i,8) to our running ticker's tickerVolume
>
>Line 70 to 74 - If statement for if cell value (i-1,1) is not equal to ticker and if cell value (i,1) is equal to ticker then cell value (i,6) is starting price
>
>>We use (i-1,1) because the ticker dataset is sorted by ticker and date. If the row value before does not match ticker then the current row value must be starting price
>
>Line 76 to 81, If statement for if cell value (i+1,1) is not equal to ticker and if cell value (i,1) is equal to ticker then cell value (i,6) is ending price
>
>>Same logic for when we did starting price but we are looking at (i+1,1) to verify ticker's ending price
>
>Line 82 - For loop to move to next i, in this case next row, and run if statements again


At this point in our subroutine, our nested loop (loop i) has run through all the rows and compiled tickerVolumes, tickerStartingPrice and tickerEndingpprices for the current ticker in the tickers(index) array. Before we move to the next ticker in our tickers index loop, we need to output the data into our table for Steve.

![Output data in worksheet](https://user-images.githubusercontent.com/96326293/148667699-f5704f12-651b-490f-9ae2-5e1eee6c795e.PNG)
>Line 84 - Declaring our output data will go into worksheet "All Stocks Analysis"
>
>Line 85 to 86 - Where data for ticker and tickerVolumes are place
>
>Line 86 - We ticker (Ending Price/ Start Price) - 1 to get the %  change over each year
>
>Line 89 - For loop to move to next Index to run for next ticker

Once the data has been added and the subroutine reaches line 89, the first loop will start again for the next index value until it reaches its end. This mean the nested loop would be run again as well. In total this subroutine will run through 15,824 rows for 11 tickers which is total of 174,064 rows ran.

## **Results**

After running the code for both years we our output data looks like this:

![All Stocks 2017](https://user-images.githubusercontent.com/96326293/148668392-5ad201cf-4579-4ce6-93d0-959983e36480.png)

![All Stocks 2018](https://user-images.githubusercontent.com/96326293/148668395-9158f2a8-0a56-4682-bfab-36010a60b0da.png)

**[Summary]**

Comparing the time elapsed from the original code vs refactored code, we saw a slight increase in time with the refactored code.
This slight increase in time is due to the output formatting section in the refactored code.
Even though our refactored code increase in time, the output formatting will benefit Steve greatly when looking at the output data.
The color coding in % return will make it easier for Steve to spot which stocks are in the positives and negatives.

**[Limitations]**

Looking back at the code, an improvement we could do is to create a variable for all the Cells.Values. This will make it easier for someone else to interpret the code if they need to edit it or if the column format changes we can easily access all the Cells.Values. Examples of how this would look like is:

![Creating variable for Cells](https://user-images.githubusercontent.com/96326293/149077180-8cf9eed8-c198-4acd-a651-987c7a7369bf.PNG)
