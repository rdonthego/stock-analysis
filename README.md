Module 2 Challenge
**OVERVIEW OF PROJECT:**
The purpose of the project for the client was to assist him to quickly analyze stock market performance over several years.  The platform was using VBA, Visual Basic for Applications within Excel.  The purpose of using VBA for this project was to access a large dataset and calculate the client’s request quickly.  The task was to refactor a code to improve it to run as quickly as possible.
****
**RESULTS:**
Project:  
I refactored code to allow a user to query for any year instead of a specific year.  Data could be added in additional worksheets for additional years and the same code can now be used.
The provided code included setting up the timer, creating the format for the output sheet and setting up or initializing the stocks to be looked at.
The project requested that a ticker index was created to determine return on stock.  That would be set to 0 as we were going to initiate a count. 
'1a) Create a ticker Index
    tickerIndex = 0
Next three array variables tied to the 12 stocks provided were defined with their data tapes using the Dim command.
'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
The looping of which data to pick up was defined next.  Looping commands tell the program to look at and do something to the data and then continue and do it through the rows defined in the command.
'2a) Create a for loop to initialize the tickerVolumes to zero.
         For i = 0 To 11
     tickerVolumes(i) = 0
'2b) Loop over all the rows in the spreadsheet.
      For j = 2 To RowCount
The next steps provided some parameters of where the loop should pick up data and what to in places to define the beginning and end.  We wanted the macro to add or increase the stock ticker volume to the current stock ticker, assigning the first row starting price and the last row with the ending process.  These were completed with If/ Then Statements.
'3a) Increase volume for current ticker
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
         
    '3b) checking that if current row is first row with with tickerIndex then closing price is tickerStartingPrices variable
     
    If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
              tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    
    End If
    '3c) checking that if current row is first row with with tickerIndex then closing price is tickerEndingprices variable
           
    If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
              tickerEndingPrices(tickerIndex) = Cells(j, 6).Value            
               
     '3d) increase the tickerIndex if next row's ticker doesn't match the previous row's ticker
          tickerIndex = tickerIndex + 1
          End If
Finally the loop command was set to calculate the percentage return over the year for the stock with the variables defined above as tickers, the tickerVolumes, and the difference (delta) between ending and starting prices.
4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
                Worksheets("AllStocksAnalysis").Activate
        
        Cells(4 + j, 1).Value = tickers(i)
        Cells(4 + j, 2).Value = tickerVolumes(i)
        Cells(4 + j, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) – 1

The remainder of the code was provided and allowed for formatting on color of output, adding percentage to the Return column, adding font variability and color return for positive and negative returns.  The colors of red and green are helpful for a client to easily see their outcomes and are an important part of macros for endusers.  

Running the macro resulted in the message box appearing asking for the year input.  I added a button so the client can easily do this in addition to the buttons previously placed during the module work.
This is the screen shot while entering the year 2017.
Data Efficiency/Improvement of Macro Outcomes:
After hitting ‘OK’ on the Message Box, I received a return of the data and a timer calculator: 
Then I redid the process with 2018 data.
 
This was the timelapse for the data to be calculated for 2017 prior to refactoring. As you can see there was a significant reduction in time after refactoring. e 
This was the timer of the 2018 data prior to refactoring. 
Stock Performance
Stocks in 2017 performed much better than in 2018.  All but one of the stocks had a positive outcome.
2018 was a much different story except for stocks ENPH and RUN with both having significant returns.
Similarly the return on the S & P was up 21.7% in 2017 and finished the year in 2018  at -4.4%.  2017 was considered by many as an “epic year” in stock trading and stocks overall did well.  DQ did well in the green stocks market that year with a great return.  Trump’s economic focus on oil and gas and industry swayed investors to ride the wave of a president with ideas to change everything and bring the economy into a level never seen before.  That had reversed after the new policies were being put into place and the plan wasn’t as effective as many people expected who were playing the market.  

**SUMMARY:**
The advantages of refactoring code are many.  In this project we were able to improve the efficiency of the macro and widen the capabilities of the analysis with some better use of variables, If/Then statements and a better loop than we originally had.  Industry application advantages are looking at an old code to improve it to find bugs, make it more understandable, and eliminating initial program “workarounds” or little codes stuck in to get the programmer by on a deadline or process he/she didn’t have time to research to do better with.  If I were using macros in my workplace I would want to relook at them on a periodic basis to tweak and clean up the inefficient “stuff” in the code.
Disadvantages of refactoring code include the lack of resources of time and money.  The code may work ok and if you don’t have time, it doesn’t get done.  Any time you go back into a program you have a good chance to affect its intention or accuracy by making a typo or inserting something that really doesn’t work.  That’s why you always run your code and make sure it works. 

 
