# Renewable Energy Stock Analysis

### Overview

Using excel VBA to create loops and more to quickly analyze stock market data with the push of a button. I'll create macros and use excels other functions to make this analysis, comprehensive and easy to digest. Initially used for one stock the code is then refactored to analyze a new set of stocks.

### Results 

I first had to refactor the original script to analyze the new set of tickers i was given. I was also to do this at a much faster and more efficient speed in order to prove that my modifications worked. In order to ensure that my new script worked I also added the following code at the end of my loop.
```
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
then when you ran the analysis and entered your prompted year, you got the time it took to run the analysis. 

*Here is the time for 2017.
<img width="229" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/107452167/199657281-57cc2861-d383-424e-8f5c-419658e084c1.png">

*Here is the time for 2018.
<img width="238" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/107452167/199657262-269e8c3d-b8e1-407c-9a4f-04d63d038ca5.png">


### Summary

While conducting this new analysis the obvious benefit of refactoring the code was that I didn't have to begin from scratch. It was much easier to see what I had already created and implement new ideas to simply have it function much quicker. The disadvantages to this however, were that any slight error could derail your thought process and take longer to find.

As for the original VBA script the only advantage again was that there was a foundation to work with and the dissadvantages were that it was much slower than after I refactored it.
