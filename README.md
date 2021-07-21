# A Study on Refactoring Code in VBA
Refactoring an Excel VBA module to determine total daily volume and total returns for 12 different stocks in 2017 and 2018.

## Overview of Project
The purpose of this project is to learn how to take already developed code and refactor it into a more useful code. To do this
we began with already developed code which looks at the 2017 and 2018 data for 12 different stocks to find their total volume and 
return value over the given year. This code is then refactored to attempt to make the code run more efficiently.

## Results
I was able to refactor the original VBA code to create a more efficient code and confirmed this code creates the same output as the
original code. Looking at the results of the time tickers for the refactored code below, it was certainly an improvement from the
original code.
When running the refactored code with the input year 2017, the time ticker showed the following result:
https://github.com/taestylobster/Stock_Returns_Analysis/blob/b4c59e9357a489a4907e6a025769838fbd3800f7/Resources/VBA_Challenge_2017.PNG
When running the refactored code with the input year 2018, the time ticker showed the following result:
https://github.com/taestylobster/Stock_Returns_Analysis/blob/b4c59e9357a489a4907e6a025769838fbd3800f7/Resources/VBA_Challenge_2018.PNG

## Summary
When it comes to refactoring code I can think of a few advantages and disadvantages after having now refactored code myself. A huge
advantage of refactoring is efficiency. By being able to make code more efficient, I can see how it can easily help reduce computing
times, reduce wear on a computer's hardware, or simply create a better experience for the end user. The biggest disadvantage, in my 
opinion, is the time it may take to refactor code. I can imagine it isn't always easy to see exactly what can be done to a code
to make it more efficient right away. It likely can take a lot of time just to determine what can be done. I can also imagine 
refactoring gets harder and harder the more times it has already been refactored, this likely plays a role on determining if taking
time to do so is worthwhile. Lastly, there are advantages to having a new person look at the code and see improvements that the
original coder(s) might not have seen. Although this can be a good thing, it can also cause a problem with the amount of time
a new person will need to just read the code and understand how it works.

As for the specific VBA code I refactored, there is a distinct advantage for the refactored code in that none of the loops are nested.
The original code had a nested loop which requires going through one loop everytime the initial loop does a single loop. In my opinion,
this made the refactored code much easier to read and understand what the code is performing. I also noticed that the run time seemed
to be much faster than the original code when I compared them side-by-side with the same computer programs open, etc. to make other 
variables as standard as possible. As for a disadvantage for the refactored code, it does require using 4 arrays versus the original
code only using 1. This creates a need to have an additional loop that resets the values within the tickerVolumes array each time the
code is run, whereas just using a variable can be easily reset each loop.
