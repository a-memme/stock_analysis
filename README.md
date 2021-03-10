# Energy Stock Analysis 
## Overview 
Refactoring code in Microsoft Visual Basic, originally created to analyze a collection of clean energy stocks. The purpose of this analysis is to first, present a clear, concise, and flexible route of analysis in analyzing a collection of energy stocks. Second, in order to be as flexible as intended, the original specific body of VBA code originally created would need to be refactored to run more efficiently, and hence, be more applicable to the potentiality of analyzing a larger set of stocks. The focus here was to rid of the nested loop originally created - although successful in analyzing the collection of stocks at hand, a nested loop would loop over the data at an exponentially larger rate with every data addition, and therefore, would be very taxing on the efficiency of the code if looking to analyze a larger dataset. Instead, the aim was to edit the code to perform the same tasks using single loops.
## Analysis and Results
### Analysis 
To begin the analysis, a base code was provided in order to provide structure to the refactoring process (See Resources Folder - challenge_starter_code). This was based on a body of code written through the modules, originally for the purpose of analyzing the twelve clean energy stocks. (See Resources Folder - modules_code). 
The first step was to create a new variable that would be used to reference a collection of additional arrays to the code, and set it to 0. This was done with reference to the natural formatting of arrays in VBA, which set the first number in their index at 0.

<img width="253" alt="tickerIndex" src="https://user-images.githubusercontent.com/79600550/110257518-5093ef80-7f6c-11eb-94a2-d48121e898a2.png">

Once the "tickerIndex" variable was made, a set of three additional arrays were then created in addition to the tickers array previously declared in the original code. These arrays would represent the values of interest when analyzing the dataset. They were then declared as long and single, respectively. 

<img width="308" alt="dim_arrays" src="https://user-images.githubusercontent.com/79600550/110257819-7ff72c00-7f6d-11eb-87f8-5ba2957056ec.png">

Second, as the tickerVolumes(12) values were to be additive by default (i.e the analysis is interested in finding the total volume of each ticker), a loop was created to initialize each element in the array to start with the value of 0. 
 
<img width="465" alt="tickerVolumes_zero" src="https://user-images.githubusercontent.com/79600550/110258196-53441400-7f6f-11eb-9e56-d5a31ddf9ebd.png">

Next, a forloop was provided, referencing the range "2 to RowCount". The variable "RowCount" represented the last row where data exists using a line of code discovered through the modules.

<img width="390" alt="RowCount" src="https://user-images.githubusercontent.com/79600550/110258687-dcf4e100-7f71-11eb-8bbe-614c4f34c48c.png">

This range would provide a loop that would search through the entire dataset in the first column (or column "A"). Here, using the three output arrays created in combination with the indexValue variable, a number of functions were performed to mirror the same tasks accomplished in the original code, but performed in a single loop. First, using the indexValue variable, a statement was created to increase the value of each element in the tickerVolumes(12) array (3a - see below). Next, conditionals were used to check whether the current cell in the loop was the first cell containing data matching the specific ticker of question. If it was, then the data in its respective "Closing Price" column (in reference to the 2017 or 2018 worksheets) would be added to the respective element in the tickerStartingPrices(12) array (3b - see below). Simiarly, the next statement also used conditionals to check whether the current cell in the loop was the last cell containing data matching the specific ticker of question. If it was, then the Closing Price value would be assigned to the respective element in the tickerEndingPrices(12) array (3c - see below). Lastly, conditionals were used again to increase the value of the variable "tickerIndex" if the following value in the loop did not equal the current value (i.e if the next row in the dataset represented a different ticker symbol - see 3d below). This conditional would allow the tickerIndex variable to become dynamic, and therefore, flow through all 12 elements in each array. Rather than using the If, ElseIf, Else formatting of conditionals, a seperate If-Then Statement was created for each conditional within the same loop.

<img width="691" alt="forloop_conditionals" src="https://user-images.githubusercontent.com/79600550/110258646-acad4280-7f71-11eb-9850-84172e383163.png">

Finally, in order to output the data, another single forloop was created, looping through numbers 0 to 11 to reference each element in the four arrays. Statements were written to output the data to the referenced worksheet "All Stocks Analysis", using the values collected in the four arrays, and filling the respective cells with the analyzed data of interest (i.e each ticker heading, total daily volume for each ticker, and yearly return % for each ticker).

<img width="572" alt="outputting_data" src="https://user-images.githubusercontent.com/79600550/110259002-6f49b480-7f73-11eb-9698-21a6754e0de1.png">

The final following blocks of code were used to change the visual formatting of the spreadsheet and were derived from the original body of code developed in the modules. To see the finished product of refactored code, reference the Resources Folder - final_refactored_code.
### Results 
In regard to the stock analysis itself, 2017 proved to be a much more fruitful year in energy stock returns vs. 2018. Eleven out of twelve stocks saw a positive yearly return, with only one company posting negative yearly returns (TERP: -7%). Quite a healthy portion of these stocks saw an extremely high yearly return, with four companies posting north of 100% gains (DQ, ENPH, FSLR, and SEDG). Total daily volume amongst all stocks hovered generally around 100-700k mark, with the exception of DQ and HASI which traded at a daily volume of 35.7k and 80.9k respecitively (see referencing below). 

<img width="292" alt="2017_allstocksanalysis" src="https://user-images.githubusercontent.com/79600550/110701336-09516d00-81bf-11eb-80e9-fd3994ca65ce.png">

2018, in contrast, saw quite the poor performance across the sector with ten out of twelve stocks yielding negative yearly returns. The poorest performing stocks were down in the -60% range, both having an average daily volume within the 100ks (DQ and JKS). The extremity of the top performers' gains, in contrast, were greater than the poor performers' losses, with 81.9% and 84% total yearly returns (ENPH and RUN respectively). Both stocks traded between the high 400k-500k's (see chart below). 

<img width="282" alt="2018_allstocksanalysis" src="https://user-images.githubusercontent.com/79600550/110701359-0f474e00-81bf-11eb-9484-486b0bfa2364.png">

In analyzing this sector between both years, one can confidently conclude that 2017 proved to be a very strong performing year across the entire  sector, while 2018 was a poor performing one. The extremity in gains/losses through these years however, indicate that this sector is potentially quite volatile, especially when considering that the average volume within individual stocks between both years generally didn't see any extreme changes. Nevertheless, further analysis regarding relative volume, overall indices' performances during the same year, etc. would need to be performed in the future to make more accurate and specific inferences.

In regard to the efficiency of code written, the purpose of refactoring code for this analysis was achieved. In the original analysis using nested loops, the code ran through the data at 1.273 seconds for the year 2017, and 1.258 seconds for the year 2018 (See Resources Folder - modulecode_2017 and modulecode_2018). Once refactored, the new code utilizing a number of arrays as well as an additional index variable to achieve the same tasks using only single loops, ran through the code at 0.203 seconds for the year 2017, as well as 0.195 seconds for the year 2018 (See Resources Folder - refactoredcode_2017 and refactoredcode_2018). With this information, one can conclude that as the new refactored code ran at a rate approaxiametly six times faster than the original code, it would be a much more efficient and applicable coding option when considering the future anaysis of any potential larger dataset.

## Summary 
Refactoring code presents both advantages and disadvantages to the process of data analysis. Generally and most importantly, editing code can provide one with a much more efficient option in performing the same tasks through code with less physical steps for the system to perform, take up less processing memory, and therefore be immensely less taxing of the system performing it. This is very important as many analytics tasks require using large datasets and complex code, and would benefit from having this type of flexibility and applicability in professional environments.
On the contrary, the downside of refactoring code is its potential visible complexity. That is, although a piece of refactored code may be more efficient systematically, it very often can be a physically longer body of code, and be more complicated or potentially more difficult to understand. In this example, the original code is slightly shorter and easier to follow as there are less variables to rely on, and its operations revolve around one core loop that accounts for all values of interest. The shortcoming in this loop as mentioned in the overview of the analysis, is that it relys on a nested loop, which tasks the system to perform an exponentially larger amount of loops through the data and thus, processes at a slower rate. With a new inclusion of data, the code will run through the data at a multiplication of the value rather than an addition if it, and would inevitably be performing redundant tasks.
Conversely, the refactored piece of code declares three new arrays and one new variable, and although more efficient, can be slightly harder to follow with its strong reliance and consistent use of these arrays/variables through the course of the code. Nevertheless, with using several seperate single forloops and strategically placed conditionals within these loops, the code gets at the data in a more direct, and thus more efficient, process.



