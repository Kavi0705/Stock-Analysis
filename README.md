# Stock Analysis with Excel+VBA

## Overview of Project
This project utilizes Microsoft Excel with VBA to analyze stock market data from 2017 & 2018 in order to inform decisions for purchasing stocks.

### Purpose
The purpose of this project is to refactor existing VBA script to allow end users to expand the analysis to many more stocks versus a few. Additionally, the aim is to determine whether refactoring the existing code allows for faster processing of the VBA script.

## Analysis and Challenges

### Stock Analysis for 2017 & 2018
In order to analyze the stock information, a VBA script was refactored in order to loop through all the data one at a time and collect the information necessary. Firstly, a ticker index was created that was used to access the correct index across 4 different arrays (tickers array, tickerVolumes, tickerStartingPrices, and tickerEnding Prices). <img width="654" alt="Stock Analysis VBA Script Step1a-3a" src="https://user-images.githubusercontent.com/93743169/148466025-3d393ec4-4dda-4ca5-a197-534436708cd3.png">

Then, a few 'for loops' and nested 'for loops' were created to initialize the tickerVolumes to zero, to loop over all the rows in the given spreadsheet (either 2017 or 2018), to generate output (Ticker, Total Daily Volume, Return), and also to format the results worksheet (named All Stocks Analysis). Within one of the for loops, If/Then statements were written to ensure rows selected and tickerIndex matched. <img width="820" alt="Stock Analysis VBA Script Step3b-4" src="https://user-images.githubusercontent.com/93743169/148466061-38ccda89-50fd-49fd-a7c4-4520fd8aab95.png">, <img width="472" alt="Stock Analysis VBA Script_Formatting" src="https://user-images.githubusercontent.com/93743169/148466104-84244542-fb4d-4014-bd80-3efe814c3450.png">

Finally, a MsgBox script was written to indicate time it takes to run the script for the year selected by the user.<img width="630" alt="Stock Analysis VBA Script_MsgBox" src="https://user-images.githubusercontent.com/93743169/148466122-bdcb6807-5290-4432-8bb6-d4db927dffee.png">

### Challenges and Difficulties Encountered
One of the challenges encountered was to ensure the Arrays were created correctly. For example, if the (12) is not included after the Array Name, then it is not an array, but rather a static variable. So, one must understand what an array truly is and what function it serves in VBA scripting.

Another challenge is to ensure the 'for loops' are nested correctly and closed correctly. One way to overcome this is to give a unique value to each for loop. In the current analysis, 'i', 'j', 'k', and 'l' were used to differentiate the 'for loops'.

## Results

Based on the analysis, it can be concluded that a majority of the stocks performed better in 2017 versus 2018, with only one stock ('TERP') having a negative return. On the other hand, in 2018 only 2 stocks performed well ('ENPH' & 'RUN'). Overall, 2017 was a better year to invest in the Stock Market.

<img width="897" alt="All Stocks Analysis_2017" src="https://user-images.githubusercontent.com/93743169/148466282-6a64d92e-c8d0-489d-9365-4e65f45032f4.png">, 

<img width="925" alt="All Stocks Analysis_2018" src="https://user-images.githubusercontent.com/93743169/148466312-6955bc74-1be7-4cc2-9d24-efd86437a004.png">.

Refactorizing the existing VBA script also allowed for it to run quicker. 

<img width="263" alt="Code Run Time for Year 2017" src="https://user-images.githubusercontent.com/93743169/148466524-cbe70baf-0c51-4f51-9009-346fcf9adc3b.png">,

<img width="256" alt="Code Run Time for Year 2018" src="https://user-images.githubusercontent.com/93743169/148466816-8e3fb2b1-be81-4220-95e8-846c63bd9e08.png">

### Advantages of Refactoring Code

Refactoring code needs to be done in small increments, which can essentially make your already existing code more efficient. You are not necessarily 'buying a new car', but rather making modifications to it. When refactoring, logical errors are easy to spot in a more structured code that utilizes for loops and if/then statements. 

### Disadvantages of Refactoring Code

Since refactoring is best done in small increments, it can take more time to write the final script. Also, you may have repetitive code that may slow down the processing time of the script. Finally, if the script is not structured and commented on appropriately, it can cause errors or utlimately impacte the outcomes.

### Original vs. Refactored Code

There are times when you may want to stick with your original code. The original code is simpler to understand and decontruct if need be. For example, if you may want to write a script on a smaller scale to ensure functionality before you rework it to incorporate more data. However, once that is accomplished, refactored data allows you to expand on the original idea for more detailed analyses. Whether you choose to use the simpler orginal code or refacator it, ensure all code is clean, well-organized, and commented on so that it can be edited or understood by others at a later time.
