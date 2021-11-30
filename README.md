# Green Stock Analysis : VBA Challenge
# Project Overview
This purpose of this project is to conduct an analysis of a portfolio of “green” stocks for a client. The client’s parents invested in one stock with ticker symbol “DQ.”  The client wants to know if this is the best stock out from a portfolio of 12 green stocks. This task is to use Excel VBA code to analyze two tables of stock data from 2017 and 2018.  For us to deliver on the clients request we must tally the total volume of stocks traded and compute the annual return for each year of available data.  Secondary requirements of the challenge require a refactoring of the code to optimize speed.

# Results of the Stock Analysis
  The analysis shows that stock with ticker DQ had a 199.4% return in 2017 but surrendered some of that with a -62.6% return in 2018.  Computing the 2-year return would give a better picture here.  Overall, all but one of these stocks showed positive returns in 2017 and then all but 2 showed negative returns in 2018.  This would indicate some correction or impact to the sector or the overall stock market.  More analysis would be required for those insights.  
  Two stocks showed positive results in both years. Tickers ENPH and RUN.  The stock, ticker symbol ENPH had impressive returns of 129.3% and 81.9%.  On face value with no company profile this would seem to be a stock worth investing in.  Stock RUN had 5.5% and 84.0% in 2017 and 2018 respectively.  This would also be a stock worth holding based on this simple analysis. (see 2017 and 2018 results tables below)
 
 <img width="208" alt="2017 Stock Analysis results" src="https://user-images.githubusercontent.com/91839403/143981143-0b960c6e-6ee7-4fee-b38a-88e1e24aed8f.png">
<img width="201" alt="2018 Stock Analysis" src="https://user-images.githubusercontent.com/91839403/143981197-742760ba-37ee-4c03-87ba-641b834afa87.png">


# Results of Refactoring
Refactoring this code from my initial working code made dramatic reductions in performance.  The initial code ran in 1,68 and 1.73 seconds.   Exceedingly long.  I was aware of of the inefficiencies in my initial code but did not fully comprehend how much slower the code would function from what is possible.
Post refactored code was significantly improved.  The efficient use of loops and thoughtful planning on how the code will function is clearly an important element in data science and design of algorithms.  The approach to doing as much as can be managed in one loop as opposed to looping through the data multiple times was the real change in refactoring that transformed the performance in this analysis.  Below are the results of the refactored code.  The 2017 analysis ran in .209s and the 2018 analysis ran in .172 seconds.


<img width="244" alt="PreRefactor 2017 Analysis" src="https://user-images.githubusercontent.com/91839403/143981251-e128787e-8d52-4c86-97bc-c80d490b0bbd.png">
<img width="230" alt="2018 Stock Analysis Timer" src="https://user-images.githubusercontent.com/91839403/143981269-ede6a77b-6ad9-48fe-8044-a47e8d28052f.png">
<img width="246" alt="PreRefactoring 2018 analysis" src="https://user-images.githubusercontent.com/91839403/143981282-dd9a1b67-dcec-4ed3-b3d9-e61c7d60bd8a.png">
<img width="230" alt="2018 Stock Analysis Timer" src="https://user-images.githubusercontent.com/91839403/143981303-24f3b14f-54bc-4965-96e3-55dce5e28508.png">


# Project Summary: Stock Analysys
The vast majority of the green stocks analyzed showed significant volatility in 2017 and 2018.  ENPH and RUN are the only two stocks that showed positive results in both years.  This is effectively only doing analysis on returns with little to no understanding of the company or how these stocks performed vs. the broader market.  That analysis would add some needed perspective and context to these results.  Given the limited information available and the analysis of these 12 stocks in 2017 and 2018 ENPH and RUN should be considered by the client for including in a portfolio.  Additionally DQ remains overall positive but did underperform in 2018 so selling now might be a consideration to preserve the gains made in 2017.

# Advantages and Disadvantages to refactoring in general
Refactoring helps clean up and optimize code for efficient operation.  It certainly is important and can help make sure the code is efficient and does not waste users time unnecessarily.  The only disadvantage, if one could this it a disadvantage, is that it takes time to perform and under time crunch that might present a challenge.  Ideally as we learn we can make more efficient code the first pass through, but I feel that will come with experience.  I would also consider investing time upfront in planning and outlining so the code can be more efficient from the start.
Refactoring has the advantage of ensuring that the code running is the most efficient it can be for a wide variety of data input. The disadvantage of refactoring is only the time 

# Advantages and Disadvantages to refactored code in Green Stock Analysis.
The main advantage to the refactored code is the significant time reduction achieved in execution.  There is a secondary advantage in the refactored code being simpler to read, understand and build upon for future enhancements or analysis.

