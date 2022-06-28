# Stock Analysis with VBA

## Overview of Project

### Purpose
In this project, the VBA script that was made in Module 2 to analyze stocks will be refactored, which is a redesign of the code to make it more efficient, take up less memory, and can potentially be made easier to read. Refactoring code also does not change the result of the code and performs all of the same tasks, but in a different way. The VBA script analyzed stocks that are in Excel worksheets named by year. The worksheets contains multiple stocks and accompanying information including price information (open/close/adjusted close/high/low) and trading volume for a given date. A table output of total daily volume and percent of return for all the stocks is made using VBA with the ability to input a year of interest in an input box to show the stock information for that year. A message box will also display the amount of time it takes for the code to run to compare the code before and after refactoring. 


## Results

![VBA_before_challenge_2017](https://user-images.githubusercontent.com/98781992/176048228-3b42074b-3b1e-4610-820e-41ba77a24a5d.png)

Figure 1. Runtime of VBA code before refactoring for the year 2017.




![VBA_before_challenge_2018](https://user-images.githubusercontent.com/98781992/176048318-90ff814a-5c4d-4ab1-9134-f162b3cca3e8.png)

Figure 2. Runtime of VBA code before refactoring for the year 2018.




![VBA_Challenge_2017](https://user-images.githubusercontent.com/98781992/176048542-5ea44f10-a1fa-4937-be0b-60eb24a189c9.png)

Figure 3. Runtime of VBA code after refactoring for the year 2017.




![VBA_Challenge_2018](https://user-images.githubusercontent.com/98781992/176048653-ff5c1c54-0916-4d3f-92fb-3ca9a641a8de.png)

Figure 4. Runtime of VBA code after refactoring for the year 2018.



**Result of refactoring (for 2017):**
- Decrease in runtime: original runtime - refactored runtime = 0.8867188 - 0.0625 = 0.8242188 seconds
- Percent decrease in runtime for code after refactoring: (decrease in runtime/original runtime) * 100 = (0.8242188/0.8867188) * 100 = 92.95%

**Result of refactoring (for 2018):**
- Decrease in runtime: original runtime - refactored runtime = 0.8671875 - 0.0625 = 0.8046875 seconds
- Percent decrease in runtime for code after refactoring: (decrease in runtime/original runtime) * 100 = (0.8046875/0.8671875) * 100 = 92.79%


## Summary

### Advantages and Disadvantages of Refactoring Code
**Advantages:** Refactoring code is useful because it can make code run faster, easier to read, and also take up less memory. This also helps with making code more scalable when potentially dealing with larger datasets.

**Disadvantages:** Refactoring code could take a significant time that would be more than the time it takes to run the original code. New bugs in the code could also be introduced. 

### Advantages and Disadvantages of the Original and Refactored VBA Script
**Advantages:** The runtime of the VBA script was significantly decreased. (See results) 

**Disadvantages:** Interpretation of the refactored code may be more difficult because a nested for loop may be more intuitive for some to understand. Also, there are multiple arrays that are made for the refactored code, which may seem more complicated than the original script that only has one array. 
