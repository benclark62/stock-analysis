# stock-analysis
## Module 2 - VBA

### Overview of Project 

Our challenge was to leverage Microsoft Excel VBA to develop Macros that will enable us to quickly analyze large datasets to determine how well a stock performed during a given period of time. Using two years' worth of daily stock information, the primary metric to idenfity high- and low-performing stocks is their in-year return.  Understanding how the frequency with which a stock is traded (Total Daily Volume) and percentage increase in price over the course of a year will be valuable inputs to making an informed investing decision.  Our initial analysis focused on a single stock - Daqo (DQ) - as our friend's parents were interested in purchasing shares.  To assist their investing decision, we broadened our analysis to include twelve total stocks.  As we expanded our analysis to include additional data, we refactored our VBA code to improve its efficiency, allowing it to process the larger dataset more quickly.

Key data provided includes **stock ticker symbol**, **trading date**, **closing price**, and **daily volume**.

![VBA_Sample_Data](https://github.com/benclark62/stock-analysis/blob/main/Resources/VBA_Challenge_SampleData.png)

The return calculation is a simple percentage increase of the stock's closing price between the first observed closing price and the last.

> Return % for stock (i) = tickerEndingPrice(i) / tickerStartingPrice(i) - 1 

### Results 
#### Stock Performance
2017 and 2018 were very different years in terms of stock performance - 2017 delivered average returns of 67.3% while 2018 returned an average loss of 8.5% across our twelve stock portfolio.  There are countless contributing factors to these results, namely the "stacking" effect of year-over-year returns that result from comparing two complete years independently.  That said, two stocks managed to deliver consecutive years of positive returns - ENPH and RUN.   ENPH was clearly the top performing stock with 129.5% returns in 2017 followed by 81.9% returns in 2018.  RUN, while not quite as strong, delivered 5.5% and 84.0% returns in 2017 and 2018, respectively.

DQ - the focus of our initial investment research - had the highest returns in 2017 with 199.4% growth.  This performance was noticed by more than our friend's parents as DQ's total volume grew by 201.4% year-over-year, the largest growth among our twelve stocks.  

Depending on our advisees' investment strategy, a two-year return could be a more valuable analysis than single year returns.  Conversely, if they are more active investors, there are shorter periods of time within 2017 and 2018 that delivered higher rates of return.

![VBA_Challenge_2year.png](https://github.com/benclark62/stock-analysis/blob/main/Resources/VBA_Challenge_2yearGraph.png)

#### Code Refactoring
As the size of the data set grew to include more stocks, it became important to refactor our *AllStocksAnalysis* macro to improve efficiency and ensure consistent performance.  There were three primary changes made to refactor the code:

- Created a *tickerIndex* to eliminate the need to looop through all tickers in individual calulations for tickerVolume, tickerEndingPrice and tickerStartingPrice. 
- Created three output arrays based on the tickerIndex.  These were previously generated as part of a larger loop through all of the tickers.
>> tickerVolumes(12)

>> tickerStartingPrices(12)

>> tickerEndingPrices(12)
- Loop and stored all tickers to intialize totalVolume = 0 first rather than keeping that calculation in a singler, larger loop. Separating this action contributed to improved efficiency for the entire macro.

Refactoring the code improved run times by ~81% compared to the original code. 

![VBA_Challenge_2018.png](https://github.com/benclark62/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
*2018 execution improved from 0.658 seconds to 0.121 seconds*

![VBA_Challenge_2017.png](https://github.com/benclark62/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
*2017 execution improved from 0.666 seconds to 0.117 seconds* 

### Summary 
#### What are the advantages or disadvantages of refactoring code?
Refactoring results in much faster script execution and more efficient code that is capable of handling larger datasets. The act of refactoring improves the user's understanding of the code's intricacies which imroves trouble-shooting and debugging and ultimately makes the user a more proficient coder over time. 

The primary disadvantage is that refactoring can be a time-consuming excercise that could prove unnecessary depending on the dataset size and the benefit created.  For example, the performance improvement for a dataset of this size was <1 second of processing time.

#### How do these pros and cons apply to refactoring the original VBA script?
There was value in refactoring the original code in this exercise because I "learned by doing".  The more time I spent manipulating the code and working through errors, the better I understood the process and will be more efficient next time.  The "con" in terms of performance improvement is certainly true - the hours spent refactoring only translated to a sub-one second improvement in processing time.  In a real world scenario, I would hestiate to commit that much time to generate nearly imperceptible performance improvements.
