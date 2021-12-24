# stock-analysis
In Module 2, I stayed with Excel and examined programing using Visual Basic for Applications (VBA). These scripts offer the ability to automate repeated functions in an Excel document, and they can be adapted from an existing script using _refractored_ code

## Overview of Project
My friend, Steve, had asked for help selecting which environmentally responsible stocks were the best to invest in. His parents had learned about DAQO Corporation in a conversation with a friend, but Steve wanted to check out the stock. He had also compiled two years of stock data on other eco-friendly corporations for me to consider.
There was so much data, spread over two years, I felt that it would be best to automate the script. Moreover, with automated VBA, Steve would be able to plug in subsequent or earlier years for a better look at trends.

## Results
I began by evaluating the stock Steve's parents had been so interested in, DAQO. 

To do this I created a For-loop to go through all the stocks in a list to find the DQ symbol. 
`    For i = rowStart To rowEnd
    
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
        
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
    
    End If`
    
Next, I programmed the file to evaluated the starting and the ending prices of DQ over the course of the year.
`    If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
    
        'set starting price
        startingPrice = Cells(i, 6).Value
        
    End If
    
    If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
        endingPrice = Cells(i, 6).Value
        
    End If
`

The results I found were significant. DQ had _fallen_ by 63% over the past year. While a company that has lost almost 2/3rds of its value may not be a lost cause (it could recover), it's not a good bet for Steve or his parents to invest a significant amount of money in. 
![DQ_Perf](https://github.com/JDittes/stock-analysis/blob/main/ActivityDQ_results.png)

### A Look at a Broader Range of Stocks, Added Time Period
When I refractored the code to add in more environmentally friendly stocks, the performance of DQ comes into greater focus. And some other stocks also show promise.

First, I ran the VBA macro for the year, 2017. Note the performance of DQ that year. While my initial review of the stock showed a 63% loss in 2018, it had _doubled_ the previous year! This shows that the stock isn't a complete stay-away, that the 2018 returns may have been a correction to the dramatic increases of 2017. This led to a net gain of 136% for the two-year period.
![DQ_2017](https://github.com/JDittes/stock-analysis/blob/main/DQ_2017.png)

This leads to a comparison with the other stocks on the index that Steve created for me to analyze.

The first thing I noticed: 2017 was an exceptional year for all the stocks on Steve's index. All but one increased in value in 2017 (TERP was the exception). And of those 10 stocks, all but 2 increased by a value of 23% or more! Look at the performance for that year.
![2017_index](https://github.com/JDittes/stock-analysis/blob/main/VBA_Challenge_2017.png)

In 2018, only two stocks gained in value, ENPH and RUN. All others lost value, including DQ. 
![2017_index](https://github.com/JDittes/stock-analysis/blob/main/VBA_Challenge_2018.png)

To help Steve with his index, I set up another page to compare the stocks' performances over the two years. I found that the average gain in 2017 had been 67% (Median 42%), while the average gain in 2018 had been -9% (Median -12%). To get a better compaison of the two years, I added a column that showed the Net Gain on the stocks, which had an average of 59%.

I found that DQ's Net Gain over the two-year period was 137%, well outpacing the average. However, ENPH gained each of the two years, and was the highest total improvement of 211%. SEDG also outperformed DQ over that time with a net 177% gain. Other stocks that outperformed the average were FSLR (+62%) and RUN (+90%). RUN, I might add, was one of the only two stocks in the index to gain value over the last year of the index.
![comparative_data](https://github.com/JDittes/stock-analysis/blob/main/comparative_index.png)

## Summary
Refractoring the AllStocksAnalysis code helped me to get down to the basics of VBA coding. I spent a lot of time trying to understand the variables I would be using in the script, and it gave me a deeper understanding of the types of variables one uses in VBA (long, single, double, integer, string, etc.)

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
