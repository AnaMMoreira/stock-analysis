# Stock Analysis

# The Use of VBA code to Perform Stock Analysis and the Application of Code Refactoring

## Overview of Project: 

In this project our good friend Steve has given us some data to analyse for his parents' investment ventures.  They are interested in investing in some green energy alternatives specifically.  They have already invested in a company called DAQO New Energy Corp which makes solar panels and has the ticker DQ.  Steve was interested in the analysis of the past years performance of their stock and had given us stock market data for the years 2017 and 2018.  With the VBA code that was prepared for Steve at the click of a button, he can analyze an entire dataset and see the DQ performance for the past 2 years. 

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/DQ_Analysis_for_2017_and_2018.png)

## Results

The table above shows the results of the code and the DQ performace sumarized in an easy to visualize and interpret format.  However, it also became clear that, eventhough DQ stock performance appeared to skyrocket and increase to close to 200% in 2017, it incured some serious losses of aproximately -63% for 2018.  Steve therefore became increasingly interested in performing the same analysis to the overall green energy market and in seeing and comparing all the stocks of the green energy alternative market. 

Although our code works well for a few stocks, it might not work well for an extensive database of stocks. Such analysis may take a long time to execute and the code would have to be modified everytime its needed.

By modifying the original code to make it more user friendly and flexible, the analysis results allows Steve to view how DQ performs overall in comparisson to the other stoks in the same market segment.  The following tables shows the resulting summary of our expanded analysis.

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/Results_AllStocks_2017.png)

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/Results_AllStocks_2018.png)

Steve can now see that, not only that his parents invested in a company that did not perform well in the past year, but also which companies also underperformed and which companies continued to perform well in the market.  The summary tables can also help him advise his parents which companies continue to do well and might be a better investment.  See Summary Table below.

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/Results_AllStocks.png)


## Refactoring Code

One way to ensure that a piece of code that has proven to be an efficient tool for a particular application continues to improve is by refactoring the code.  In lamest terms, refactoring means to rewrite existing text to improve readability, reusability, or structure without intentionally detracting from its meaning. Similarly, in the context of coding, refactoring code makes it more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future use without changing meanning or behavior.  This makes Refactoring a standard practice in any data analysis field and is often a starting point at a new job.



## Comparison of original and refactored VBA script

In this challenge, refactoring the previous code to loop through all the data one time in order to collect the same information for all different stocks (Tickers) made sense in order to generate the tables previously shown. The following comparison points out the differences in code and how it affected the scripts performance.  The original code was harwired to perform a particular year and only for the ticker in question "DQ".  Although we added a button to make it more user friendly it originally did not prompt the user to select a Year and the code would have to be manually changed to get each year.  It was repetitive, long and harder to follow. 
The following image provides a side by side comparison.  As can be observed, the original code in Subroutine DQAnalysis() is much longer and looks more "messy" and therefore harder to follow than the refactored code in Subroutine AllStocksAnalysisRefactored().  

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/code_length.png

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/variable_compare.png)

Further the original code uses many more variables than the refractored code wich intuitively means that it uses more memory resources in order to yield a limited set of output (only one ticker and only one year at a time) the image below shows how each code works.

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/compare_for_loops_process.png)

Note that on the image on the right (AllStocksAnalysis()), after the user is propted to input a year, each ticker is passed through the same simple loop for each ticker in an array.  This process is not only simpler on paper, but also simpler to solve, and  should process faster and use less memory resources.  To ensure that in fact the processes of the refactored code is faster and more efficient a timer was added and the result printed on screen for both the original code and the refactored code.  The following images shows the resulting run times for each code and what the user input window looks like.   

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/OriginalStockCode_2018_Timer.png)

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/UserPrompt_2017Input_example.pgn)

and 

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/2017AllStocks_Timer.png) 

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/UserPrompt_2018Input_example.png) 

and 

! [] (https://github.com/AnaMMoreira/stock-analysis/blob/main/Resources/2018AllStocks_Timer.png)


Notice that for the original code it took 0.0625 seconds to run one year worth of data for one ticker alone.  And for each year, the AllStocksAnalysisRefactored Subroutine took ~ 0.715 seconds to process and print the stats for all stocks.  At first hand it looks like the Refactored code may not be much more efficient in runtime because the results sugest that the run time of the original code for each ticker would add up to 0.75 seconds.  But considering that the user would have to add or modify the code each time a new ticker analysis was needed and also format the resulting table (which was also automated) for the original code it can be assumed that it would be much more time consuming.


## Summary:
It can thus be concluded that Refactoring code is most advantageous.  I can definately automate analysis that are run repetitavely and used often on a daily basis.  It acomplishes this by producing reliable results that are fast, uses less process and storage resources, and is user friendly.  Overall this would result in more acurate performance, reducing the need to QAQC results as often, decreasing production costs and increase productivity.  Furthermore, the develoment and polishing of existing internal coorporate or standard tools, makes the upgrade or Refactoring of existing code faster to develop and apply. 




