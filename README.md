# stock-analysis
Excel and VBA Application
## Overview of the Project
### Purpose
<p align="justify"> The purpose of this project is to analyze different green energy stocks in order to help Steve and his parents who want to invest in a particular one named *DAQO*. By analyzing all the green energy stocks Steve is going to be able to have a bigger picture and advise his parents in a better way. 
The analysis of the project will be done downloading all the green stocks information for the years 2017 and 2018, and then using VBA and Excel to automate tasks. By using this programming language Steve will be able to run the analysis for each year any time he wants. </p>

## Results
There were a series of steps taken to make the analysis of the green stocks. First, the outline of the code was downloaded and pasted into the VBA_Challenge file. Then, a ticker Index was created as well as 3 output arrays called: *tickerVolumes*, *tickerStartingPrices*, and *tickerEndingPrices*. 

<img width="309" alt="1" src="https://user-images.githubusercontent.com/111388644/191076860-47488b3f-a83c-4b04-9bfc-9289181af4b2.png">

<p align="justify"> After that, two **for loops** were created. One to initialize the *tickerVolumes* to 0, and the other one to loop over all the rows in the spread sheet. To make the second loop the code from the original script was copied. Later, a code was written to increase the current *tickerVolumes*, which was taken from the hint in the bootcamp spot. </p>

<img width="503" alt="2" src="https://user-images.githubusercontent.com/111388644/191076913-32f87e6d-71c2-4c8b-9279-3183d7901edb.png">

Following that, 2 **If-then** statements were made to assign the starting and the closing time.

<img width="481" alt="3" src="https://user-images.githubusercontent.com/111388644/191076976-0c45fa1f-1b19-4b7d-a6c7-b6babeb8c627.png">

Then, as the image below shows, a script that increased the ticker Index if the ticker of the current row does not match the next ticker. And finally, a loop that loops through the arrays was made. 

<img width="618" alt="4" src="https://user-images.githubusercontent.com/111388644/190714039-7e32015d-ef30-4731-9e57-0bc8609b748a.png">

<img width="683" alt="5" src="https://user-images.githubusercontent.com/111388644/191077037-13feeeb0-7904-4d58-b041-ed142bb6607e.png">

<p align="justify"> There is a lot of difference between the stock performance in 2017 and in 2018. Based on the images displayed below we can conclude that 2017 was a better year regarding the stock performance since only one of the 12 had a negative return. Or in other words, the TERP stock had a loss of 7.21%. The rest of the stocks had positive returns. The highest return was 199.45% of the DQ stock. </p>

<img width="311" alt="analysis 2017" src="https://user-images.githubusercontent.com/111388644/190714136-7cbff9d5-2f60-461a-ba98-bc363b1adf33.png">

On the other hand, the 2018 stock performance was not the best for most of the green stocks. In this year only two stocks had a positive return while the rest of them had a negative one. 

<img width="280" alt="analysis 2018" src="https://user-images.githubusercontent.com/111388644/190714164-42e642a7-a3b2-47c1-ade2-e5db6de4fb3f.png">

The execution time for the original script for 2017 and 2018 was of 1.42 and 1.18 seconds respectively. This time was slower than the one ran in the refactored script.   

<img width="280" alt="run time for module 2 2017" src="https://user-images.githubusercontent.com/111388644/190714666-d01e9d7a-cfe5-4d6c-8414-3ad7dcb8745f.png">

<img width="306" alt="VBA_Challenge_2018 png" src="https://user-images.githubusercontent.com/111388644/190714312-a887cc67-9ffb-4a30-abf3-0d809d9c3250.png">

The execution time for the refactored script was 0.24 for 2017 and 0.26 for 2018.

<img width="314" alt="run time for refactored 2017" src="https://user-images.githubusercontent.com/111388644/190714438-00aa1e4e-a4c7-4a8e-bb2f-24557d2e9025.png">

<img width="301" alt="VBA_Challenge_2018 refactored png" src="https://user-images.githubusercontent.com/111388644/190714478-463c5712-6ef8-4cb6-9738-1cb45063e08c.png">

## Summary
### Advantages and Disadvantages of refactoring code
<p align="justify"> There are certain advantages when it comes to refactoring code such as not having to write everything from the start. By using code that has already been written you can save time by only editing some parts of the code. Also, you do not have to create the outline or the structure of the code again since it is already organized the only thing there is to do is change it to run the subroutine that you want. 
Of course, there are some disadvantages like having to debug the code. This may take a lot of time because there can be certain errors that are easy to miss when copying another code. </p>

### Pros and Cons Applied to refactoring the original script
<p align="justify"> When refactoring the original script, I had a lot of problems when it came to debugging because I kept finding new errors and the code did not run. First it took me a lot of time to realize which were the errors I was committing, and after I struggled fixing them. Another disadvantage I found was that not every single part of the original code worked, so I had to get the help pf a tutor that explain to me why some parts had to be removed. 
Although the part I just described was difficult, there were some pros I found. One of them was that I was able to start to work on the code right away since there were already the steps to follow. 
Without doubt this assignment was very challenging but, at the end when the code ran the way it was meant to, it was very fulfilling. </p>

