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

The results I found were significant. DQ had _fallen_ by 63% over the past year. While a company that has lost almost 2/3rds of its value may not be a lost cause (it could recover), it's not a good bet for Steve or his parents to invest a significant amount of money in. ![DQ Results] (https://github.com/JDittes/stock-analysis/blob/main/ActivityDQ_results.png)

### A Look at a Broader Range of Stocks, Added Time Period
When I refractored the code to add in more environmentally friendly stocks, the performance of DQ comes into greater focus. And some other stocks also show promise.

First, I ran the 

The analysis is well described with screenshots and code (4 pt).
## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
