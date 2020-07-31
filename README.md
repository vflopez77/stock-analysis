Module 2 Challenge Analysis

# Refactoring Stock Analysis Code and Its Effect on Performance

## Overview 
The Stock Analysis code used a number of loops to produce the desired results.  By refactoring the code, we may be able to improve the performance.  To this end, we have rewritten the VBA script to use arrays to store the data for each ticker requiring only one loop through the data.  We will measure the performance of the script using timers to see the effect of the refactoring.