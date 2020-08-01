Module 2 Challenge Analysis

# Refactoring Stock Analysis Code and Its Effect on Performance

## Overview and Purpose
The Stock Analysis code used a number of loops to produce the desired results.  By refactoring the code, we may be able to improve its performance.  To this end, we have rewritten the VBA script to use arrays to store the data for each ticker requiring only one loop through the data.  We will measure the performance of the script using timers to see the effect of the refactoring.

### Results
The refactored code produced the same data results as the original code: 

<img src=/Resources/Refactored_Output_2017.png></img>
<img src=/Resources/Refactored_Output_2018.png></img>

However, the refactored code performed about <b>7 times faster.</b>

Original and Refactored Timings for 2017:

<img src=/Resources/Unrefactored_Timing_2017.png></img>
<img src=/Resources/VBA_Challenge_2017.png></img>

Original and Refactored Timings for 2018:

<img src=/Resources/Unrefactored_Timing_2018.png></img>
<img src=/Resources/VBA_Challenge_2018.png></img>

### Analysis of Code Refactoring
The original script looped through all the data rows 12 times - once for every ticker.

The code looked like this:

    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
    'Loop through rows in the data.
        Worksheets(yearValue).Activate
            
        For j = 2 To RowCount
            
    'Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
    ...
        Next j
    ...
    Next i

The refactored script loops only once through the dataset, storing the data for each ticker in arrays by using the tickerIndex variable:

    tickerIndex = 0
    ...
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
    ...  
             tickerIndex = tickerIndex + 1
    ...           
    Next i

By using the arrays tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices() to store the summary data and scanning the records only once, we have more efficiently used computer resources and significantly improved execution speed.

### Summary
It is important in programming to write the most efficent code possible in order to most effectively use the available computing resources of CPU, memory, and I/O.  Sometimes it is necessary, because of time constraints or the need to prove that a solution is possible, to use "brute force", inefficient programming methods.  However, we should always strive to ultimately produce the most compact and elegant code by refactoring the code, testing each change along the way.  Besides the performance gains, there are other important advantages to rewriting code to make it as good as possible.  One is readibility, so that the code can be more easily maintained by others.  Another is that by reusing pieces of code without rewriting it in multiple places, we avoid transcript errors and having to make the same changes in multiple places.  On the other hand, refactoring code that already works is time-consuming and costly, and code that is too compact (and not commented enough) may be difficult for others the understand.  Overall, refactoring is worth the effort because it makes best use of computer resources, and leaves a useful code base for further development.

In this particular application refactoring the code, as described above, has yielded a real increase in performance.  While a fraction of a second may not seem particularly significant, it is crucial if we are to use larger data sets.  In the real world, the number of records could easily be in the millions instead of thousands.  Additionally, the new structure makes it easily extensible so that multiple year sheets could be processed sequentially.  While working with multiple arrays may be a little more difficult to conceptualize, refactoring the code to make it as efficient allows us to scale the application and handle larger datasets.



