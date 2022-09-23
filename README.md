# VBA-challenge
This is for Homework2

## Explanation of the code in VBA_Code_Homework2

The 2 sub routines Unique() and PriceChange() are designed to run on one worksheet at a time and iterate through the total number of sheets.

The subroutine Unique() iterates through Column 1 and write down each unique ticker symbol encountered to the column with header 'Ticker'

Once all the unique tickers are noted, the subroutine PriceChange() will calculate the Yearly change, percent change, total stock volume in the first half of the script(including the conditional formatting to color the cells).

Later in the second half of the script, it calculates the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

I compared my results with what is provided in the resources folder and it matches for the values. However, I used "Insert empty column function" to match with the screenprint "hard solution"
