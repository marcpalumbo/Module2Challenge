**# VBA_Challenge_StockChallenge

## Overview

For this challenge, we were tasked with refactoring, or improving upon, previously written code to make it more efficient and include more elements than what was covered throughout the instruction portion of the module. Refactoring on its own would have presented its own challenge, but refactoring and adding elements to it introduced a different level of challenging concepts that challenged one's overall understanding of VBA, and entry level concepts of programming.

The end product is an interactive VBA program that allows for a user to easily which year they would like to run a thorough stock analysis on for the 12 stocks presented that Steve's parents are looking to invest in. The program will allow for Steve's parents to evaluate each of the 12 stocks simultaneously over the two-year period of 2017 and 2018. Below will be a summary of what our analysis shows in regards to best and worst stocks to invest in, as well as how much we were able to improve upon the efficiency of the program. 

## 2017 Stock Performance

The performance of the 12 stocks analyzed in 2017 was far superior than that of 2018. This is immediately made evident by the formatting done on all stocks that yielded a positive return at years end. 2017's analysis shows that all but one stock yielded a positive return. Comparatively, 2018 showed that only two stocks yielded positive return. TERP ended up being the only stock with negative performances in both years. Below is the performance of all stocks analyzed in 2017: 
![2017 Stock Performance and Run Time](https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/VBA_Challenge_2017.png)

If Steve's parents are looking into stocks to invest in, it would be in their best interest to consider running the analysis on both years before making concrete decisions. They would be making a severe monetary risk if they evaluated all of these stocks based on their performance in 2017.

## 2018 Stock Performance 

The performance of all stocks analyzed in 2018 is grim to say the least. All but two stocks are in the negatives in terms of yearlong returns. This return was calculated using the simple calculation of subtracting the stock price at the end of the by the price at the beginning of the year. One thing that Steve's parents may want to take into consideration is that most of the stocks were bound to not perform as well as they did in 2017. The general idea of regression to the mean would say that these stocks were not likely to outperform themselves in consecutive years. However, even with that in consideration the two stocks that did show consecutive years of positive returns may be their best option for investment. At least, if it was my money, those would be what draws my attention first. Listed below are the results and run time for our code when running analysis on 2018.
![2018 Stock Performance and Run Time](https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/VBA_Challenge_2018.png)

## Original Code vs Refactored Code Performance
After refactoring our code, we were able to shave off over 1 full second on our run time WHILE adding additional variables and elements to analyze. This was done by changing the order in which we ran our code through VBA. In AllStocksAnalysis(), We began formatting cells in between different parts where we were still analyzing information. This May have contributed to the vast differences in run time. Below I will include the screenshots of the run times between the two subroutines. 
### Run Time on Module Before Refactoring (2017)
![2017 Run Time](https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/First_Runtime_for_%202017.png) 
### Run Time on Module After Refactoring (2017)
![2017 Stock Performance and Run Time](https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/VBA_Challenge_2017.png)

### Run Time on Module Before Refactoring(2018)
![2018 Run Time[(https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/First_Runtime%20for_2018.png)
### Run Time on Module Before Refactoring (2018)
![2018 Run Time and Stock Performance](https://github.com/marcpalumbo/Module2Challenge/blob/main/Resources/VBA_Challenge_2018.png)
## challenges and how they were overcome. 

The most challenging part of this module was figuring out how to get the provided code in the instruction to work. This took me the longest as most of the code was transferrable to the challenge portion. I had asked for assistance using MyBCS and the tutor explained to me that I needed to include the array, either "(i) or "(12)"each time I was referencing the array in the module. In hindsight, I realize this should have been obvious, but everything else I had written would have worked perfectly had I don’t this earlier. One of my main challenges in learning programming will definitely continue to be the meticulous attention to detail that it requires when referencing variables or arrays. I must do a better job at remembering that I need to take the approach of programming as me literally talking to a computer and telling it what to do step by step. In day-to-day life, we are so used to our devices filling in the blanks and realizing what we are trying to say in a few keywords like when we google something. This is not reality when using programming languages. I am the one creating the ease of use for another user, and I must do the detailed work so that it can be user friendly.  






