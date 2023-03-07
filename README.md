# VBA-challenge

# VBA Homework: The VBA of Wall Street

# Hello! The instructions did not really say what to include in my read me file, so I am going to explain my thought process while writing my VBA script.

# In my script, first I had to define the values that I needed to store. I needed to store:
- The ticker value
- The total stock value
- The year open price for each ticker
- The year close price for each ticker
- The yearly change in price for each ticker
- The percent change in price for each ticker

# I set the initial total stock volume as 0 so that we could add to it later.

# I also knew that I would need to display all of my results in a summary table, and this would have to start in the second column, so I added that into my code as well.

# I then determined the last row of the sheet, so that later on I could tell the code to run until the last row.

# I ran the script from the second row (the first row of data) until the last row. 

# I told the script to look at the row below the current row, and if the ticker value was not the same, to do the following:
- Store the ticker value for the row I am currently in - it is the final row of that ticker
- Store the year close price from that row - it is December 31st of that year so it is the closing price
- Add to the total stock volume by adding the stock volume from that row to the previous stock volume total
- Display the ticker and total stock volume in the summary table
- Calculate and store the yearly change by subtracting the year close price from the year open price
- Display the yearly change in the summary table
- Calculate and store the percent change by dividing yearly change by year open price and then multiplying by 100
- Display the percent change in the summary table
- Add 1 to the summary table row so that we can get ready to start displaying data for the next ticker
- Set the total stock volume to 0 so we can start adding up the stock volume for the next ticker

# If the row below the current row had the same ticker value, then I told the code to:
- Add to the total stock volume by adding the stock volume from that row to the previous stock volume total
- If the ticker was the first data entry of the year (entered on Jan 2), then store the year open price

# For the next part of my script, I ran another for loop to search for certain values:
- First I went to the percent change column and searched for whichever ticker had the maximum value
- I put this ticker and its' value as the greatest % increase
- Then I searched the percent change column for whichever ticker had the minimum value
- I put this ticker and its' value as the greatest % decrease
- Then I searched the total stock volume column to find the maximum value
- I put this ticker and it's value as the greatest stock volume

# Finally I had to loop this code through all sheets of the data:
- For this I used a for statement to tell the code to run through every ws in Worksheets
- Any time I discussed a location of a cell, I had to specify to do it in the current worksheet
- Thus, I had to put ws before any time I wrote Range or Cells so that it would know to go to these locations in the current sheet

I hope this explains my VBA script to you. Enjoy!
