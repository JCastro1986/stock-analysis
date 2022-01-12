# stock-analysis
Stack Analysis for Steve
# VBA of Wall Street
Challenge  #2. Create an analysis to help Steve's parents better invest their money in the best stock.
# Overview of Project
I will be helping my friend Steve who just graduated with his Finance degree. I will create an analysis to help Steve's parents to invest in the better stock, they are passionate about green energy and want to be Steve's first client. I will help Steve by creating a flexible macro that runs analysis in multiple stocks, in order for Steve's parents to have the best possible information before deciding which company to invest.My analysis will include the following information for the years 2017 and 2018.
        *Ticker*
        *Total Daily Volume*
        *Return*
This project is to help Louis understand how campaigns are going based on their launch dates and their fundings goals.
## Results
Analyzing a handful of green energy stocks for 2017 and 2018 in addition to DQ stocks (Steve's parents go to stock) using VBA code (the programming language that can interact with excel). The usage of VBA reduces the chance of accidents and errors.
Using the current data we have created a detailed and easy-to-read summary of the different stock available (Hydro, Wind, Geothermical and Bio companies) and now we can visualize all different stock by year.
We can notice the following:
    - 2017: All stocks but TERP were giving positive returns. 
        * Stocks with top return were (DQ, SEDG and ENPH)
        * The only stocks with negative return was TERP.
            ![2017 Results](P:\DATA ANALYST - TEC\Challenges\Module 2\Resources\All Stocks (2017).png) 
            ![All Stocks (2017)](https://user-images.githubusercontent.com/95668609/149052983-592039a9-ed3e-4dea-8a3b-983d2d47a5a9.png)
    - 2018: Only two stocks were giving positive returns.
        * The only stocks with positive returns were ENPH and RUN, (RUN with the highest return %).
        * Stocks with top negative returns were (DQ, JKS and SPWR).
            ![2018 Results](P:\DATA ANALYST - TEC\Challenges\Module 2\Resources\All Stocks (2018).png)
            ![All Stocks (2018)](https://user-images.githubusercontent.com/95668609/149052998-eaf034fb-f4ab-4277-9138-617bfdbecdc4.png)
The key components of my coding were:
    - 'Initialize array of all tickers
    Dim tickers(11) As String
    - '1a) Create a ticker Index
    tickerIndex = 0
    - '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single  
    - ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
    - '3d Increase the tickerIndex.
    tickerIndex = tickerIndex + 1
I conclude that although DQ was giving positive returns in 2017 it is not doing that great in 2018, rather it is giving negative returns.
I recommend investing in **RUN**, which is the stock giving the best returns in 2018. 
*In case that Steve's parents want to diversify their investments their money should only go to **RUN** and **ENPH**, both giving positive returns in 2018.*
## Summary
- What are the advantages or disadvantages of refactoring code?
1. Advantages:
    - We clean up our current code.
    - We avoid magic numbers.
    - We can turn our code more efficient by using new tools such as:
        *An Index.
        *Arrays.
        *Conditionals.
2. Disadvantages:
    - We can, by mistake, make our code run into an error.
        *We can mistakenly erase something our code had when runnign running, thus creating bugs.
    - We can change the results of our code.
        *It is important to check our code often make sure we have not to change the behavior of our code.
        (This happened to me, I by mistake made some updates and before checking I save the updates, later I realized I did something that change the result of my code and it was very hard to find and correct it).
- How do these pros and cons apply to refactoring the original VBA script?
1. Pros:
    - Pros were great since I already had the code running properly, I was just looking to make it better, more efficient, and cleaner.
    - I was able to create the missing "Index" I had pending and thus reducing the amount of information within my code.
    - I removed information that was not adding any value, this made my file look cleaner and more readable.
            * Before:   '3a) Initialize variables for starting price and ending price
                        Dim startingPrice As Single
                        Dim endingPrice As Single
            * After:    '1b) Create three output arrays
                        Dim tickerVolumes(12) As Long
                        Dim tickerStartingPrices(12) As Single
                        Dim tickerEndingPrices(12) As Single
2. Cons:
    - While trying to clean my code, I y mistake of removing important information that made my code change behavior resulting in a different result.
        *I was able to check my code before submitting it but it was too late, so I needed to go back and review every piece of coding in order to find the mistaken information.
            * Before:   '3a) Initialize variables for starting price and ending price
                        Dim startingPrice As Single
                        Dim endingPrice As Single
            * After:    '1b) Create three output arrays
                        Dim tickerVolumes(12) As Long
                        Dim tickerStartingPrices(12) As Long
                        Dim tickerEndingPrices(12) As Long
                        *The usage of Long variables in the tickerStartingPrices and tickerEndingPrices resulted in different results of my code.*
