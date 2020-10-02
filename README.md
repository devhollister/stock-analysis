# An Analysis of Selected Stocks and the Benefits of Refactoring Scripts

## Overview of Project
The following is an analysis of the benefits of refactoring a VBA script created to analyze the performance of 12 selected green energy stocks for the years 2017 and 2018. The script was initially written to provide insights to the client, Steve, who has been tasked with investing in stocks for his parents, who are passionate about green energy and are particularly interested in the performance of a company called DAQO New Energy Corp (ticker “DQ”). The script was then refactored to allow the client to easily and efficiently scale his stock analysis to include the performance of any or all stocks on the market for any given year. To determine whether the refactoring improved the original script, run times were calculated for compared for both. 

## Results

### Stock Perfomance

In total, 12 green energy stocks were analyzed over two years, 2017 and 2018. In this analysis, the stocks are referenced by their ticker, an abbreviation used to identify stocks on the market. The indices used to measure stock performance were the total daily volume (a measure of how often the stock was traded), and the yearly return (the percentage difference in stock price from the beginning to the end of the year). Negative yearly returns (stocks that decreased in value) are highlighted in red while positive returns are highlighted in green. The results are as follows: 

![2017 Stock Performance](Resources/all_stocks_analysis_2017.png)

![2018 Stock Performance](Resources/all_stocks_analysis_2018.png)

As shown above, all stocks except for TERP gained value over the course of 2017. Over the course of 2018, however, all stocks lost value, the exceptions being ENPH and RUN, which continued to increase in value. Between those two stocks, ENPH saw a modestly lower return in 2018 than 2017 (81.9% vs 129.5%) while RUN saw a markedly higher return in 2018 than 2017 (84.0% vs 5.5%). The price of DQ stock, which Steve's parents initially were interested in investing in, fell 62.6% over the course of 2018, despite seeing an increase in total daily volume (35,796,200 in 2018 vs 107,873,900 in 2017) and having traded very positively in 2018 (199.4%). This may suggest that further analysis may be necessary to determine whether total daily volume is a useful indicator of green energy stock performance. The only stock to have a negative return over both years was TERP. 

### Script Comparison

## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt
