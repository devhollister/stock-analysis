# An Analysis of Selected Stocks and the Benefits of Refactoring Scripts

## Overview of Project
The following is an analysis of the benefits of refactoring a VBA script created to analyze the performance of 12 selected green energy stocks for the years 2017 and 2018. The script was initially written to provide insights to the client, Steve, who has been tasked with investing in stocks for his parents, who are passionate about green energy and are particularly interested in the performance of a company called DAQO New Energy Corp (ticker “DQ”). The script was then refactored to allow the client to easily and efficiently scale his stock analysis to include the performance of any or all stocks on the market for any given year. To determine whether the refactoring improved the original script, run times were calculated for compared for both. 

## Results

### Script Comparison

The primary objective in refactoring the VBA script was to allow it to better handle much larger datasets than the one included in this analysis. To do this, attention was given to removing processes that slow a script when processing large datasets, which in this case was namely nested for loops. 

### Stock Perfomance

In total, 12 green energy stocks were analyzed over two years, 2017 and 2018. In this analysis, the stocks are referenced by their ticker, an abbreviation used to identify stocks on the market. The indices used to measure stock performance were the total daily volume (a measure of how often the stock was traded), and the yearly return (the percentage difference in stock price from the beginning to the end of the year). Negative yearly returns (stocks that decreased in value) are highlighted in red while positive returns are highlighted in green. The results are as follows: 

![2017 Stock Performance](Resources/all_stocks_analysis_2017.png) ![2018 Stock Performance](Resources/all_stocks_analysis_2018.png)

As shown above, all stocks except for TERP gained value over the course of 2017. Over the course of 2018, however, all stocks lost value, the exceptions being ENPH and RUN, which continued to increase in value. Between those two stocks, ENPH saw a modestly lower return in 2018 than 2017 (81.9% vs 129.5%) while RUN saw a markedly higher return in 2018 than 2017 (84.0% vs 5.5%). The price of DQ stock, which Steve's parents initially were interested in investing in, fell 62.6% over the course of 2018, despite having traded very positively in 2018 (199.4%) and seeing an increase in total daily volume between the two years (107,873,900 in 2017 vs 35,796,200 in 2018). This may suggest that further analysis may be necessary to determine whether total daily volume is a useful indicator of green energy stock performance. The only stock to have a negative return over both years was TERP. 

## Summary

In general, refactoring makes a script more efficient, more adaptable and easily edited, and more readable and understandable. Refactoring however can be time consuming, and as a script grows in complexity, the time and potential roadblocks involved in refactoring may exceed its benefits and even the project budget. For this VBA script in particular, refactoring enabled to run THIS MUCH MORE QUICKLY by eliminating nested for loops, enabling the code to be read more clearly from top to bottom without forcing the reader to cycle through a complex loop structure.

As for the stocks performance aspect of this analysis, it appears that stock prices in the green energy field are falling in general, and though further analysis on larger and more up to date data sets should be made to confirm this, it may not be wise to invest in green energy stocks at the moment. DQ in particular should be avoided due to its especially negative performance in 2018. 
