# VBA-Challenge
# Getting this thing started was a bit tricky, but I'm satisfied with the results overall (regardless of how clunky the code may be)

# I utilized some of the code we learned in class to separate each row by ticker value and assign each unique ticker to be displayed in a separate column

# To find the total volume of each stock, I simply summed all of the volumes within the rows for each ticker

# Finding the yearly change fofr each stock in each worsheet was, for me, the most diifcult section of the assignment without a doubt
# I took advantage of the fact that each worksheet gave the same number of row for each stock, and that each stock's last date was the 31st of December
# To find the yearly change, I wrote the code to identify the last four values of the date string('right' function). If they happend to be 1231, that cell's I-value would be subtracted by the appropriate number of rows for that worksheet to find the first row of data for that year
# I also made use of the 'left' function so that I could make separate condictionals for each year that would also run simultaneously
#From there I performed the necessary calculations to find both the yearly change and percent change
# I refrained from multiplying by the percentage value by 100 within the code itslef so that I would be able to format as percentage in excel without any issues
# My originally worked only on the 2020 sheet, this turned out to be because my 'left' function value should have been four for '18 and '19 instead of five like it was for '20


# I used a worksheet function to find the required maximums and minimums and then used if statements to dictate where a vlaue should be displayed should it match one of the values found by the worksheet funtions

# The conditional formatting function in excel was used to change the fill for each cell in the 'yearly change' column based on the cell value
