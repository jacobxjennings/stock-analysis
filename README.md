## Overview of Project
In this project, I use Excel and VBA to analyze historical stock data.

## Purpose
Previously, my client requested assistance in building a automated script to help analyze stock data for his parents investment fund. I was given a dataset with historical stock data for 12 stocks in a two year period. I wrote a script using VBA to quickly be able to analyze datasets and output specific stock metrics. In the future, this allows for new data to be easily imported and analyzed without having to build a worksheet for every dataset. Upon completion, I was instructed to refactor my code to improve the efficiancy of the script. 

## Inital Analysis and Challenges

Initally, I believed my script ran pretty well. However, after further analysis I realized that I could improve on a couple of things to allow the script to scale more efficiently. For example, after making a few changes to the script, shown below, I reduced the time the code takes to run by .0986 seconds. 

![Time Result Before Refactoring](https://github.com/jacobxjennings/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Time.PNG)
![Time Result After Refactoring](https://github.com/jacobxjennings/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored_Time.PNG)

Listed below, are an example of the inital code for increasing the totalVolume. These two do the exact same thing but the second uses arrays to process the data. 

```
If Cells(J, 1).Value = ticker Then
            
        totalVolume = totalVolume + Cells(J, 8).Value

End If
```
```
If Cells(XX, 1).Value = tickers(tickerIndex) Then
    
        tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(XX, 8).Value

End If
```

Since the goal of refactoring is to just improve on the process of the same code, the outputs should and do look the exact same. This is show below: 

![Analysis Output Before Refactoring](https://github.com/jacobxjennings/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Output.PNG)
![Analysis Output After Refactoring](https://github.com/jacobxjennings/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored_Output.PNG)

## Summary
Refactored code, in general, is faster and more efficent to compute. This is especially expressed when working with huge datasets that could take minutes to analyze. When refactoring code one must realize that readability is deminished; therefore, they must be extra cautious to include descriptive comments, as well as organized lines/whitespace. 
