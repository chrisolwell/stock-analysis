# stock-analysis
Using VBA to analyze stock performance.

## Project Overview
Steve approached the writer to help disuade his family from making what he believed to be a bad investment by analyzing stock data to show the rate of return for 2017 and 2018. We later expanded the scope of the project in a way that theoretically allows us to analyze any stock for any year, so we refactored out code so that it would be able to process larger datasets. 
### Our Analysis
Although we also tracked trading volume for the stocks, our analysis is relatively unsophisticated: we sought only to learn the whether the stock's price increased or decreased over the course of a calendar year, as well the percentage of the increase or the decrease, which is a simple division problem if you know the price of the stock when markets opened on the first day of the year and when they closed on the last day of the year. Unfortunately, we didn't have those specific data point. What we have instead is the opening price and closing price for each day of the year, as well as the daily trading volume, so we needed to manipulate the data into something we can work with.
#### Our Method
We wrote code to perform our analysis in the modules, and here in the challenge we refactored it so it would work faster. To juice up our code, we used indexing, which allows the code to keep track of what it's already counted, instead of counting it again everytime the program loops through our code. We initiallized our ticker index with the code *tickerIndex = 0.*
As we mentioned, refactored is intended to work faster. Our client liked our stock analysis so much he wanted to use it to analyze any stock, which means our original code wouldn't have performed as well because it was intended to analyze only 12 stocks. Our refactored code indexes data it has already looked at, which makes it more efficient; as we see in the screenshots that follow, our code ran in less than one-tenth of a second.

<img width="261" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/4724180/149678368-713a5ce1-5f0f-44a5-8177-a255b390d7be.png">
<img width="261" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/4724180/149678370-a0af00ea-b290-488d-8c22-5a9970131bd3.png">

I didn't screenshot the original code's performance and I have not been instructed to do so as part of this assignment, so I'm not going to. Trust me though: the refactored code performed considerably faster than the original.
#### Advantages and disadvantages of refactoring 
I have explained already the benefit of refactoring code: it's faster. You might not miss two-tenths of one second when you're only looping through 3,000-or so rows of data, but it will be important with larger datasets. Now, the primary disdvantage to refactoring VBA code is that **it is extrememly confusing.** Refactoring this code was very frustrating, and even now, after hours and hours spent writing it, I am not sure I could explain why it works, or how I got it to work, and I hope I never have to do it again.
