# An Analysis of Stock Returns
Performing analysis on green stocks' trading volumes and their year end returns to help Steve and his parents determine investment strategies.

## Overview of Project

### Purpose
Steve just graduated with his finanace degree and for his first job, his parents are going to be his clients. His parents are enthusiatic about green energy and have decided to invest all their money into Daqo New Energy Corp (_$DQ_) because they met at a Dairy Queen. They haven't done more research than that. Steve is looking into _$DQ_ as well as other green stocks to help his parents diversify their portfolio. The goal of this analysis is to help Steve and his parents.

## Results

### Stock Performance
Based on the analysis, the green stocks included in this analysis, largely had positive returns in 2017 and largely had negative returns in 2018. 

![2017_stock_results](https://user-images.githubusercontent.com/92613639/139624645-30b29014-419c-46b0-b41e-275faa263006.png)
![2018_stock_results](https://user-images.githubusercontent.com/92613639/139624654-d0c9499c-33d0-46bf-978e-adf40d13d54b.png)

It appears as though $ENPH and $RUN stand out in the 2018 output as the two stocks that were able to have positive results despite the rest of the green stocks posting negative results. This may mean that they had a particularly good year for some reason or something they do helps them diffentiate themselves and allows them to break away from the way the market moves, unlike the rest of the analyzed stocks. This is purely speculatio, and further understanding of the stocks would be required prior to providing an investment strategy suggestion.

Suggestions to Steve's parents would also depend on how risk averse they are and what they would like to see in terms of results from their investments. We see some stocks have smaller returns and losses (such as _$AY_ and _$TERP_) and that may be a safer investment if they are looking to be careful with their investments. But if they decide they are looking to go all in on the green energy movement and want to be a part of the market movers, they may choose to go more high risk, high reward with stocks that swing more intensely (such as _$ENPH_, _$SEDG_, or even _$DQ_). Looking at overall market trends of the green energy space may be important as well. It may be in Steve's best interest to find other industries even - maybe other causes they care about - to truly help his parents diversify their investments. 

### Refactored Script vs. Original
The first thing I noticed about the code was how fast the refactored code was able to run as compared to the original script. In the refactored script, we run through the data and create and array (see code below) with the relevant information prior to printing it all out into the results table. Arrays tend to be more time efficient than loops. 

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Additionally, the refactored script includes the formatting code, which the original script did not, so not only is the refactored code longer and include more steps, but it also runs faster.

## Summary

### Advantages and Disadvantages of Refactoring Code
Advantages:
- The refactored script runs faster due to increased efficiency and logic.
- The code is easier to read for future users as it generally has fewer steps.
- Refactoring someone else's code can be a good way to learn the functionality and intent of the code.

Disadvantages:
- Refactoring code can take time and energy to provide the same results as the original script, so depending on the length of the code or amount of data being analyzed, it may or may not be worth it.

### How do these pros and cons apply to refactoring the original VBA script?
- One short fall of the way this code is written is that this code relies on the data being sorted by date. Ideally this code would be written in a way that would allow us to find the starting and ending prices by looking at the earliest and latest date for each ticker and using those values so we were not dependent on the way the data is sorted.
- 
