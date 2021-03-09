# Energy Stock Analysis 
## Overview 
Refactoring code in Microsoft Visual Basic, originally created to analyze a collection of clean energy stocks. The purpose of this analysis was to improve the specific body of VBA code in order to make it run more efficiently, and hence, make it more applicable to the potentiality of analyzing a larger set of stocks. The focus here was to rid of the nested loop originally created - although successful in analyzing the collection of stocks at hand, a nested loop would loop over the data at an exponentially larger rate with every data addition, and therefore, would be very taxing on the efficiency of the code if looking to analyze a larger dataset. Instead, the aim was to edit the code to perform the same tasks using single loops.
## Analysis and Results
### Analysis 
To begin the analysis, a base code was provided in order to provide structure to the refactoring process (See Resources Folder - challenge_starter_code). This was based on a body of code written through the modules, originally for the purpose of analyzing the twelve clean energy stocks. (See Resources Folder - modules_code). 
The first step was to create a new variable that would be used to reference a collection of additional arrays to the code, and set it to 0. This was done with reference to the natural formatting of arrays in VBA, which set the first number in their index at 0.

<img width="253" alt="tickerIndex" src="https://user-images.githubusercontent.com/79600550/110257518-5093ef80-7f6c-11eb-94a2-d48121e898a2.png">

Once the "tickerIndex" variable was made, a set of three additional arrays were then created in addition to the tickers array previously stated in the original code. These arrays would represent the values of interest when analyzing the dataset. They were then declared as long and single, respectively. 

<img width="308" alt="dim_arrays" src="https://user-images.githubusercontent.com/79600550/110257819-7ff72c00-7f6d-11eb-87f8-5ba2957056ec.png">

Second, as the tickerVolumes(12) values were to be additive by default (i.e the analysis is interested in finding the total volume of each ticker), a loop was created to initialize each element in the array to start with the value of 0. 
 
<img width="465" alt="tickerVolumes_zero" src="https://user-images.githubusercontent.com/79600550/110258196-53441400-7f6f-11eb-9e56-d5a31ddf9ebd.png">

Next, a forloop was provided, referencing the range "2 to RowCount". The variable "RowCount" represented the last row where data exists using a line of code discovered through the modules.

<img width="390" alt="RowCount" src="https://user-images.githubusercontent.com/79600550/110258687-dcf4e100-7f71-11eb-8bbe-614c4f34c48c.png">

This range would provide a loop that would search through the entire dataset in the first column (or column "A"). Here, using the three output arrays created in combination with the indexValue variable, a number of functions were performed to mirror the same tasks accomplished in the original code, but performed in a single loop. First, using the indexValue variable, a statement was created to increase the value of each element in the tickerVolumes(12) array (3a - see below). Next, conditionals were used to check whether the current cell in the loop was the first cell containing data regarding the specific ticker of question. If it was, then the data in its respective "Closing Price" column (in reference to the 2017 or 2018 worksheets) would be added to the respective element in the tickerStartingPrices(12) array (3b - see below). Simiarly, the next statement also used conditionals to check whether the current cell in the loop was the last cell containing data regarding the specific ticker of question. If it was, then the Closing Price value would be assigned to the respective element in the tickerEndingPrices(12) array (3c - see below). Lastly, conditionals were used again to increase the value of the variable "tickerIndex" if the following value in the loop did not equal the current value (i.e if the next row in the dataset represented a different ticker symbol - see 3d below). This conditional would allow the tickerIndex variable to become dynamic, and therefore flow through all 12 elements in each array. Rather than using the If, ElseIf, Else formatting of conditionals, a seperate If-Then Statement was created for each conditional within the same loop.

<img width="691" alt="forloop_conditionals" src="https://user-images.githubusercontent.com/79600550/110258646-acad4280-7f71-11eb-9850-84172e383163.png">

Finally, in order to output the data, another single forloop was created, looping through numbers 0 to 11 to reference each element in the four arrays. Statements were written to output the data to the referenced worksheet "All Stocks Analysis", using the values collected in the four arrays, and filling the respective cells with the analyzed data of interest (i.e each ticker heading, total daily volume for each ticker, and yearly return % for each ticker).

<img width="572" alt="outputting_data" src="https://user-images.githubusercontent.com/79600550/110259002-6f49b480-7f73-11eb-9698-21a6754e0de1.png">

The final following blocks of code were used to change the visual formatting of the spreadsheet and were derived from the original body of code developed in the modules. To see the finished product of refactored code, reference the Resources Folder - final_refactored_code.
### Results 
In regard to the stock analysis itself, 2017 proved to be a much more fruitful year in energy stock returns vs. 2018. Eleven out of 12 stocks saw a positive yearly return, with only one company (TERP) posting negative yearly returns (-7%). Quite a healthy portion of these stocks saw a much higher-than-average yearly return, with four companies posting north of 100% gains (DQ, ENPH, FSLR, and SEDG). Total daily volume amongst all stocks hovered generally around 100-700k, with the exception of DQ and HASI which traded at a daily volume of 35.7k and 80.9k respecitively (see referencing below). 
2018, in contrast, saw quite the poor performance across the sector with ten out of twelve stocks yielding negative yearly returns. The poorest performing stocks were down in the 60% range, both settling within the 100k daily volume area (DQ and JKS). The top performers' gains here, however, outshot the worst performers with 81.9% and 84% total yearly return (ENPH and RUN respectively). Both stocks traded between the high 400k's to 500k.

In regard to the code efficiency, the purpose for refactoring the code for this analysis was achieved. In the original analysis using nested loops, the code ran through the data at 1.273 seconds for the year 2017, and 1.258 seconds for the year 2018 (See Resources Folder - modulecode_2017 and modulecode_2018). Once refactored, the new code utilizing a number of arrays as well as an additional index variable to achieve the same tasks using only single loops, ran through the code at 0.203 seconds for the year 2017, as well as 0.195 seconds for the year 2018 (See Resources Folder - refactoredcode_2017 and refactoredcode_2018).

