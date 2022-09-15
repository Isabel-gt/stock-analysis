# stock-analysis
Excel and VBA Application
## Overview of the Project
### Purpose
The purpose of this project is to analyze different green energy stocks in order to help Steve and his parents who want to invest in a particular one named *DAQO*. By analyzing all the green energy stocks Steve is going to be able to have a bigger picture and advise his parents in a better way. 
The analysis of the project will be done downloading all the green stocks information for the years 2017 and 2018, and then using VBA and Excel to automate tasks. By using this programming language Steve will be able to run the analysis for each year any time he wants. 

## Results
There were a series of steps taken to make the analysis of the green stocks. First, the outline of the code was downloaded and pasted into the VBA_Challenge file. Then, a ticker Index was created as well as 3 output arrays called: *tickerVolumes*, *tickerStartingPrices*, and *tickerEndingPrices. 

ss

After that, two **for loops** were created. One to initialize the *tickerVolumes* to 0, and the other one to loop over all the rows in the spread sheet. To make the second loop the code from the original script was copied. Later, a code was written to increase the current *tickerVolumes*, which was taken from the hint in the bootcamp spot. Following that, 2 **If-then** statements were made to assign the starting and the closing time. 

ss

Then, as the image below shows, a script that increased the ticker Index if the ticker of the current row does not match the next ticker. And finally, a loop that loops through the arrays was made. 

ss 

There is a lot of difference between the stock performance in 2017 and in 2018. Based on the images displayed below we can conclude that 2017 was a better year regarding the stock performance since only one of the 12 had a negative return. Or in other words, the TERP stock had a loss of 7.21%. The rest of the stocks had positive returns. The highest return was 199.45% of the DQ stock. 

Ss

On the other hand, the 2018 stock performance was not the best for most of the green stocks. In this year only two stocks had a positive return while the rest of them had a negative one. 

ss

The execution time for the original script for 2017 and 2018 was of 1.29 and 1.32 seconds respectively. This time was slower than the one ran in the refactored script that ended up being (VALUE for 2017 and VALUE for 2018). 

Ss

Ss


## Summary
### Advantages and Disadvantages of refactoring code
There are certain advantages when it comes to refactoring code such as not having to write everything from the start. By using code that has already been written you can save time by only editing some parts of the code. Also, you do not have to create the outline or the structure of the code again since it is already organized the only thing there is to do is change it to run the subroutine that you want. 
Of course, there are some disadvantages like having to debug the code. This may take a lot of time because there can be certain errors that are easy to miss when copying another code. 

### Pros and Cons Applied to refactoring the original script
When refactoring the original script, I had a lot of problems when it came to debugging because I kept finding new errors and the code did not run. First it took me a lot of time to realize which were the errors I was committing, and after I struggled fixing them. Although the part I just described was difficult, there were some pros I found. One of them was that I was able to start to work on the code right away since there were al ready the steps to follow. 
With out doubt this assignment was very challenging (BUT SI LA COMPLETAS PON ALGO COOL). 

