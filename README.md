# VBA of Wallstreet

## Overview of Project

### We were given a dataset to analysis stock performance of several green energy companies spanning two years. The data was compiled in Excel and we used Visual Basic to automate the analysis process. The macro we wrote in VBA scanned data for twelve different stock tickers and collected starting ticker prices, ending ticker prices and the total volume for each ticker. That data was then entered into a separate worksheet to help visualize the performance of each company. We were challenged to improve the performance of the macro by refactoring the original code. After the refactoring, the macro completed the analysis 500% faster than the original code.


## Results

### Our goal was to analysis stock data from 2017 and 2018 of twelve green energy companies. Our VBA macro pulled from the dataset and complied the ticker, volume and performance of each company on a new worksheet. We were tasked to analysis the ticker DQ from the dataset and compare its performance with the performance of the other eleven companies. In 2017, DQ performed the best of all the other companies with a return of 199.4%. DQ also had the lowest volume when compared to the other companies which could be due to several factors, like having a higher ticker price or being a new unknown company. Overall, 2017 was a good year for performance will eleven out of twelve companies showing a positive return. Performance decreased substantially in 2018 with only two companies showing positive returns. DQ saw a loss of 62.6% in 2018, which was the largest loss when compared to the other eleven companies in the dataset. DQ’s volume increased by 301% YoY which was the largest increase in the sector. 

### Our original code ran at 0.535 seconds on average. Refactoring shaved off over 0.4 seconds bringing our runtime down to 0.101 seconds. The main performance enhancement came from using arrays to compile data as the macro traversed through the dataset, instead of having the macro run through the dataset for each ticker. The original code set a variable called ‘ticker’ and would scan the data looking for that ticker to collect and transfer to the new worksheet, which meant that the macro would have to go through the dataset twelve times to collect all the data. The refactored code collected data for all the tickers as it ran through the code only once.

![2017 After Refactoring](https://github.com/mgochis/VBA_Challenge/blob/8d963a4ba65d8ef2bbb417b674aae95c44741a60/Resources/VBA_Challenge_2017.png)
![2018 After Refactoring](https://github.com/mgochis/VBA_Challenge/blob/8d963a4ba65d8ef2bbb417b674aae95c44741a60/Resources/VBA_Challenge_2018.png)

## Summary

- What are the advantages or disadvantages of refactoring code? Refactoring code enhances the performance of the code which could dramatically decrease processing times when working with large datasets. Refactoring also helps understand and learn code that is unfamiliar to you.

- 1.	How do these pros and cons apply to refactoring the original VBA script? One disadvantage to refactoring code is the time spent refactoring may not end up having a substantial increase in performance.
