Module 2 Challenge Analysis

# Refactoring Stock Analysis Code and Its Effect on Performance

## Overview and Purpose
The Stock Analysis code used a number of loops to produce the desired results.  By refactoring the code, we may be able to improve its performance.  To this end, we have rewritten the VBA script to use arrays to store the data for each ticker requiring only one loop through the data.  We will measure the performance of the script using timers to see the effect of the refactoring.  Refactoring the code to make it as efficient as possible will allow us to scale the application and handle larger datasets.

### Results
The refactored code produced the same results as the original code. 

Here is the output:

<img src=/Resources/Refactored_Output_2017.png></img>
<img src=/Resources/Refactored_Output_2018.png></img>

The timings for the refactored code were about 7 times faster than the original code.

Original and Refactored Timings for 2017:

<img src=/Resources/Unrefactored_Timing_2017.png></img>
<img src=/Resources/VBA_Challenge_2017.png></img>

Original and Refactored Timings for 2018:

<img src=/Resources/Unrefactored_Timing_2018.png></img>
<img src=/Resources/VBA_Challenge_2018.png></img>

Let's take a look at how this performance improvement was achieved.

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

The refactored script loops only once through the dataset, storing the data for each ticker in arrays:

    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
    ...  
             tickerIndex = tickerIndex + 1
    ...           
    Next i



