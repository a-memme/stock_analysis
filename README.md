# Energy Stock Analysis 
## Overview 
Refactoring code in VBA originally written to analyze a collection of clean energy stocks. The purpose of this analysis is to improve this specific body of VBA code in order to make it run more efficiently, and hence, make it more applicable to the potentiality of analyzing a larger set of stocks in the future. The focus here was to rid of the nested loop originally created - although successful in analyzing the collection of stocks at hand, a nested loop would increase at an exponentially larger rate with each potential addition to the data, and therefore be very taxing on the efficiency of the code if looking to analyze a larger dataset. Instead, the aim was to edit the code to perform the same tasks using single loops.
## Analysis and Results
### Analysis 
To begin the analysis, a base code was provided in order to provide structure to the refactoring process (See Resources Folder - challenge_starter_code). This code was based on a body of code written through the modules, originally for the purpose of analyzing the twelve clean energy stocks. (See Resources Folder - modules_code). 
The first step was to create a new variable that would be used to reference a collection of additional arrays to the code, and set it to 0. This was done pertaining to the natural formatting of arrays in VBA, which set the first number in their index at 0.

<img width="253" alt="tickerIndex" src="https://user-images.githubusercontent.com/79600550/110257518-5093ef80-7f6c-11eb-94a2-d48121e898a2.png">

Once the "tickerIndex" variable was created, a set of three additional arrays were created (in addition to the tickers(12) array), representing the values of interest when analyzing the dataset. These arrays were declared as long and single, respectively. 

<img width="308" alt="dim_arrays" src="https://user-images.githubusercontent.com/79600550/110257819-7ff72c00-7f6d-11eb-87f8-5ba2957056ec.png">

Second, as the tickerVolumes(12) values were to be additive by default (i.e the analysis is interested in finding the total volume of each ticker), a loop was created to initialize each index in the array to start at 0. 
 
<img width="465" alt="tickerVolumes_zero" src="https://user-images.githubusercontent.com/79600550/110258196-53441400-7f6f-11eb-9e56-d5a31ddf9ebd.png">

Next, a forloop was provided, referencing the range "2 to RowCount". The variable "RowCount" represented the last row where data exists using a line of code discovered through the modules.

<img width="390" alt="RowCount" src="https://user-images.githubusercontent.com/79600550/110258687-dcf4e100-7f71-11eb-8bbe-614c4f34c48c.png">

This range would provide a loop that would search through the entire dataset in the first column (or column "A"). Here, using the three output arrays created, in combination with the indexValue variable, a number of functions were performed to mimick the tasks the original code accomplished. First, using the indexValue variable, a statement was created to increase the value of each index in the tickerVolumes(12) array (3a - see below). Next, conditionals were used to check whether the current cell in the loop was the beginning in the set of data regarding a specific ticker (3b - see below). If it was, then the data in its respective "Closing Price" column (in reference to the 2017 or 2018 worksheets) would be added to the respective index in the tickerStartingPrices(12) array. Simiarly, the next statement also used conditionals to check whether the current cell in the loop was the end of the set of data regarding the ticker in question (3c - see below). If it was, then the Closing Price value would be assigned to the respective index in the tickerEndingPrices(12) array (3d - see below). Lastly, conditionals were used again to increase the value of the variable "tickerIndex" if the following value in the loop did not equal the current value (i.e if the next row in the dataset represented a different ticker symbol). Rather than using the If, ElseIf, Else formatting of conditionals, a seperate If-Then Statement was created for each conditional within the same loop.


in order to a) increase the volume of each ticker in the index; b) use conditionals to check whether the current cell in the loop is the beginning of the set of data regarding the specific ticker of question; c) use conditionals to check whether the current cell in the loop is the end of the set of data regarding the ticker of question; and d) use conditionals once again to increase the value of the variable "tickerIndex" if the following value in the loop did not equal the current value (i.e if the next row in the dataset represented a different ticker symbol - see below). Rather than using the If, ElseIf, Else formatting of conditionals, a seperate If-Then Statement was created for each conditional within the same loop.

<img width="691" alt="forloop_conditionals" src="https://user-images.githubusercontent.com/79600550/110258646-acad4280-7f71-11eb-9850-84172e383163.png">

The pieces of code above 


Finally, in order to output the data, another single forloop was created, looping through numbers 0 to 11 to reference each index in the four arrays. Statements were written to output the data to the referenced worksheet "All Stocks Analysis", filling the respective cells with the analyzed data of interest (i.e each ticker heading, total daily volume for each ticker, and yearly return % for each ticker).

<img width="572" alt="outputting_data" src="https://user-images.githubusercontent.com/79600550/110259002-6f49b480-7f73-11eb-9698-21a6754e0de1.png">

The final blocks of code following were used to visually format the spreadsheet and were derived from the original body of code written through the modules. To see the finished product of refactored code reference the Resources Folder - final_refactored_code.
