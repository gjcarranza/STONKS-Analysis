# STONKS-Analysis
---
Purpose

The purpose of this analysis is to verify which stocks are more viable for Steve's parents to invest in as a means to diversify their portfolio rather than have their full investment in only one stock (DAQO). 

---

Results

The method used to analyze such data is VBA in excel. Through various stages of coding implemented to create a primary worksheet that is more user-friendly that resulted in the depiction of twelve known stocks, their total volume, and their return rate for the respective years of 2017 and 2018. A button was inserted into the worksheet to ease the transition between both of the years provided. Two modules were developed, each with their respective coding that share a common goal: to be able to create a user-friendly button that displays one of the year values' values. The first module (Module 1) had some trouble wtih coding. One of the main issues that arose was the button itself. The user must enter the year at least three times before the timer feature pops up. The other issue is not being able to permanently display the year that the user is asking (i.e. if the user asks for 2018 stocks, the user will get the results for 2018 but will have the window that asks "What year would you like to run your analysis on?" appear twice more, and then revert back to showing the user 2017 stocks with the header title row changed to "All Stocks (2017)"). Another challenge were the "End If" statements. As seen below, the "End If" statement for my code is "End" because when th traditional statement was implemented, it read as an error. The same issues is evident when using the refactoring coding. The second module (AllStocksAnalysisRefractored) has the following coding: 

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
            tickers(0) = "AY"
            tickers(1) = "CSIQ"
            tickers(2) = "DQ"
            tickers(3) = "ENPH"
            tickers(4) = "FSLR"
            tickers(5) = "HASI"
            tickers(6) = "JKS"
            tickers(7) = "RUN"
            tickers(8) = "SEDG"
            tickers(9) = "SPWR"
            tickers(10) = "TERP"
            tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        tickerIndex = 0
        
    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0

           Next i
           
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End
                
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
                End
                
        
        '3c) check if the current row is the last row with the selected ticker
         
             If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
             
                End
                
            

        '3d) Increase the tickerIndex.
            
             If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
            
                End
                
    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For k = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
                Cells(4 + k, 1).Value = tickers(k)
                Cells(4 + k, 2).Value = tickerVolumes(k)
                Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
                
        
    Next k
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

---

Despite the issue the timers for both Module 1 and Module 2 run roughly around the same time. I have noticed that the timer is dependent on how fast the user types in the year they are attempting to search. The following images provided below displays the times for both modules for the year '2017':

![image](https://user-images.githubusercontent.com/92961267/140470034-9282e30c-5111-4353-86fb-436ce0464301.png)
   (Module 1 timer)

![image](https://user-images.githubusercontent.com/92961267/140470148-d124dc41-3d01-4fa4-b123-eec490e2f887.png)
   (Module 2 timer)

---

The following data tables provided the values of the twelve stocks' total daily volumes and return rates for the years 2017 and 2018:

![all stocks 2017](https://user-images.githubusercontent.com/92961267/140467372-4b8728a7-8091-4e21-b0c0-a1ed47b837ba.PNG) 

![all stocks 2018](https://user-images.githubusercontent.com/92961267/140467375-5a4c3483-a0d7-4556-b377-d087f62c0511.PNG)

---

Recommendation:

The recommendation upon which stocks for Steve's parents to invest in and diversify their portfolio are stocks ENPH and RUN. Stocks ENPH and RUN have managed to maintain their return rates thus investing in these two stocks may plateau at times with little risk of losing money. If Steve's parents are interested in playing a game of 'Risk', VSLR is a stock to keep an eye one. Although VSLR has decreased only slightly there is still a chance of gaining with this stock over time. Investing in VSLR is to proceed with caution. My recommendation is to do more research on VSLR and only invest the minimum amount.

---

Summary

The advantages of refactoring coding is that the coding is: cleaner, more organized, somewhat more automated (e.g. less coding is entered), viewer-friendly for reading, easier maintance; however the disadvantage of refactoring coding is the amount of time spent to create it and having to consistently test the coding to check if it is working or not. In comparison to the using the original scripts, the advantages of using the original script is that it is easier to manipulate than refactoring coding and its repeatability (i.e. can tell the computer to repeat something as many times as you'd like). However there are significantly more disadvantages in using the original script such as the original script is more prone to errors; needs consistent testing; having to manually edit said errors; must tell the computer exactly in what order to process the commands; and close-to-illegible coding.
