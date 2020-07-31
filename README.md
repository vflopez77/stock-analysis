Module 2 Challenge Analysis

# Refactoring Stock Analysis Code and Its Effect on Performance

## Overview and Purpose
The Stock Analysis code used a number of loops to produce the desired results.  By refactoring the code, we may be able to improve its performance.  To this end, we have rewritten the VBA script to use arrays to store the data for each ticker requiring only one loop through the data.  We will measure the performance of the script using timers to see the effect of the refactoring.  Refactoring the code to make it as efficient as possible will allow us to scale the application and handle larger datasets.

### Results
The refactored code produced the same results as the original code. 

Here are the results for 2017:

<img src=/Resources/Refactored_Output_2017.png></img>

