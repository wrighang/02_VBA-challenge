# VBA-challenge

## Requirements

## Retrieval of Data
The script loops through one quarter of stock data and reads/stores all of the following values from each row:
- Ticker symbol
- Volume of stock
- Open price
- Close price

## Column Creation
On the same worksheet as the raw data, or on a new worksheet, all columns were correctly created for:
- Ticker symbol
- Total stock volume
- Quarterly change ($)
- Percent change

## Conditional Formatting
Conditional formatting is applied correctly and appropriately to:
- Quarterly change column
- Percent change column

## Calculated Values
All three of the following values are calculated correctly and displayed in the output:
- Greatest % Increase
- Greatest % Decrease
- Greatest Total Volume

## Looping Across Worksheet
The VBA script can run on all sheets successfully.

## GitHub/GitLab Submission
All three of the following are uploaded to GitHub/GitLab:
- Screenshots of the results
- Separate VBA script files
- README file
-------------------------------------------------

Included in repository: 
1. Screenshots of the results 
2. Separate VBA script files
3. README file
4. EXTRA - xlsm file

Reset Button- Andrew Lane provided this code in a study group we had with several classmates over the weekend and I used it after I had completed the assignment as it was very helpful to reset using a macro versus manually deleting the results

Greatest% Increase/Greatest Decrease- I researched how to write a code that would look up the maximum and minimum values. Initially, I struggled with the formula for calculating the greatest total volume and reached out to Andrew for advice, assuming the issue was related to the min/max formula and worked on that for awhile which you can see in my code the various attemps i was making with the max code for the volume. However, it turned out that the problem was with how I was resetting the volume for the total stock volume calculation. I had originally placed the volume = 0 reset after the loop, but while adjusting things to test different outputs, I mistakenly moved it before the loop without realizing how this affected the results. I didn’t notice the incorrect output until much later.

Message Box- Since the code takes some time to run, I added a message box at the end to display a notification when the code has finished running.
