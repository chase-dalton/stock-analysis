# An Analysis of Stock Data using VBA

## Overview of Project
The purpose of this project was two-fold. Not only were we tasked with using VBA to help Steve analyze stock data, but to refactor the code in order to increase computational efficiency so the code is able to handle larger datasets. Our goal with the VBA code was to provide Steve with an Excel worksheet that would allow him to effectively and efficiently analyze various stock performance by returning the total daily volume and the return for any given year. 

## Results
The findings for 2017 showed increases in returns for the majority of stocks that were analyzed. The only stock that didn't have an increase in returns was TERP. 

![2017](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/2017_Results.png)

On the other hand, 2018 showed a general decrease for those same stocks, with only ENPH and RUN showing an increase in returns.  

![2018](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/2018_Results.png)

It wasn't only the stock values that changed, but execution times as well. By refactoring the code, we were able to improve the speed of our macro.

Here were the times for generating 2017's and 2018's code before refactoring:

![2017 Old Time](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Old_Time.png) ![2018 Old Time](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Old_Time.png) 

And here were the times for generating 2017's and 2018's code after the refactoring:

![2017 Time](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) ![2018 Time](https://github.com/typicalchazz/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

There was over a 96% decrease in execution time for both 2017 and 2018.

## Summary
Overall, our intervention with Steve was a great success. By refactoring our previous VBA, we were able to provide Steve with a tool that will allow him to analyze his usual data in a dynamic and timely manner.

Refactoring code has its advantages and disadvantages. One of the advantages is that with refactoring, you already have a solid framework to work with. The functionality of the code is already there; the code present already executes certain actions that yield certain, expected results. You don't have to start from scratch and create a brand new product. With a solid framework, all you attention can be focused on improvement as opposed to creation. 

In this example, we already had the logic in place that for each ticker, the starting and ending price was found, the total volume was summed, and that information was output to a sheet with some logic to output a return percentage and how long the code took to execute. We also hade the luxury of knowing that whatever our new code did, it had to output the same values as the original, but faster. This makes testing simple: if those criteria were met, the refactoring was a success.

One of the disadvantages of refactoring is that you are violating the idiom "If it ain't broke, don't fix it". While version control and other proper development protocols help mitigate issues regarding breaking previously working code, there is always room for error. In the world of coding there are things such as time, money, and energy to take into consideration when developing versus refactoring code. Sometimes the opportunity cost is too great and you'd be better off moving to other projects as opposed to spending your time trying to make something that works a bit better. While the new macro increased performance speed by over 96%, it made a 3-4 second process happen under 1 second. While the new code is an improvement and should be seen as truly "better", 3-4 seconds is not a lot of time to get the answers you want in the first place, and the time spent making this slight improvement could have been spent developing something more useful.