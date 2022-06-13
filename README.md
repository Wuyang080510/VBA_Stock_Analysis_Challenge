# VBA_Stock_Analysis_Challenge
Use VBA script to automate stock analysis

# Stock Analysis with VBA
## Overview of Project 

In this project, I used VBA to help Steve analyze green energy stocks so that he could give his parents stock investment advice. To do this, I need to loop through all the data and collect the total daily volume and annual return rate for each stock in 2017 and 2018 respectively. 

### Purpose
The purpose of this analysis is to use VBA script to find an efficient way to automate data extracing, calculating, and formatting, wchich is usually conducted in Excel manually. To make the code more efficient, I refactored the code to get all the data I need in only one loop.

## Results
### Original VBA Code
In the original VBA code, I used two loops to scan all the data in the file, returned the expected output in the "All Stocks Analysis" worksheet, and recorded the code running time with a timer. 
The code running time for 2017 is 1 second. The code running time for 2018 is 0.95 seconds. 
    
    Sub AllStocksAnalysis()
      'create start time and end time variables
      Dim startTime As Single
      Dim endTime As Single
    
      Worksheets("All Stocks Analysis").Activate
      yearValue = InputBox("What year would you like to run the analysis on?")
      startTime = Timer
    
      'add title in cell A1
      Range("A1").Value = "All Stocks(" + yearValue + ")"
    
      'add 3 columns with headers
      Cells(3, 1).Value = "Ticker"
      Cells(3, 2).Value = "Total Daily Volume"
      Cells(3, 3).Value = "Return"
    
      Worksheets("DQ Analysis").Activate
      'add title in cell A1
      Range("A1").Value = "DAQO (ticker:DQ)"
    
      'add columns with headers
      Cells(3, 1).Value = "Ticker"
      Cells(3, 2).Value = "Year"
      Cells(3, 3).Value = "Total Daily Volume"
      Cells(3, 4).Value = "Return"
    
      'create an array that holds 12 elements
      Dim tickers(11) As String

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
    
      'initialize variables for starting price and ending price
      Dim startingPrice As Single
      Dim endingPrice As Single

      'Activate data worksheet
      Worksheets(yearValue).Activate

      'Get number of rows to loop
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      'loop through tickers
      For i = 0 To 11
          ticker = tickers(i)
          totalVolume = 0

          'loop through rows in the data
          Worksheets(yearValue).Activate
          For j = 2 To RowCount
            
              If Cells(j, 1).Value = ticker Then
                  totalVolume = totalVolume + Cells(j, 8).Value
              End If

              If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                  startingPrice = Cells(j, 6).Value
              End If

              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                  endingPrice = Cells(j, 6).Value
              End If
            
          Next j
        
          'output data to All Stocks Analysis
          Worksheets("All Stocks Analysis").Activate    
          Cells(4 + i, 1) = ticker
          Cells(4 + i, 2) = totalVolume
          Cells(4 + i, 3) = (endingPrice / startingPrice) - 1
        
          'output data to DQ Analysis
          Worksheets("DQ Analysis").Activate
          If ticker = tickers(2) Then
              Cells(4, 1).Value = ticker
              Cells(4, 2).Value = yearValue
              Cells(4, 3).Value = totalVolume
              Cells(4, 4).Value = endingPrice / startingPrice - 1
          End If

      Next i
    
      endTime = Timer
      MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

      'worksheet format
      Worksheets("All Stocks Analysis").Activate
      Range("A3:C3").Font.Bold = True
      Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
      Range("A3:C3").HorizontalAlignment = xlCenter
      Range("A3:C3").Font.Color = vbBlack
      Range("B4:B15").NumberFormat = "#,##0"
      Range("C4:C15").NumberFormat = "0.0%"
      Columns("B").AutoFit

      'apply conditional formating to the return column
      dataStart = 4
      dataEnd = 15

      For k = dataStart To dataEnd
          If Cells(k, 3) > 0 Then
              'Color the cell green
              Cells(k, 3).Interior.Color = vbGreen
          ElseIf Cells(k, 3) < 0 Then
              'Color the cell red
              Cells(k, 3).Interior.Color = vbRed
          Else:
              'Clear the cell color
              Cells(k, 3).Interior.Color = xlNone
          End If
      Next k
    End Sub

![Code_Run_Time_2017_Original](https://user-images.githubusercontent.com/106395288/173252279-00428fb0-4862-4f22-ada3-feab6acda2fd.png) 
![Code_Run_Time_2018_Origunal](https://user-images.githubusercontent.com/106395288/173252297-561bda5a-27c6-4b9d-bcdc-09ddb478ad01.png)

### Refactored VBA Code
In the refactored VBA code, I modified the original code, reduced the loops to one, and presented all the information in the "All Stocks Analysis" worksheet.  
The code running time for 2017 is 0.21 seconds. And the code running time for 2018 is 0.17 seconds. The refactored code is much faster than the original VBA code I wrote for the stock analysis. 

    Sub AllStocksAnalysisRefactored()
      Dim startTime As Single
      Dim endTime  As Single

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
      Dim tickers(11) As String
    
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
      Worksheets(yearValue).Activate

      'Get the number of rows to loop over
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      '1a) Create a ticker Index
      Dim tickerIndex As Single
          tickerIndex = 0

      '1b) Create three output arrays
      Dim tickerVolumes(11) As Long
      Dim tickerStartingPrices(11) As Single
      Dim tickerEndingPrices(11) As Single

      ''2a) Create a for loop to initialize the tickerVolumes to zero.
      tickerVolumes(tickerIndex) = 0
      
      ''2b) Loop over all the rows in the spreadsheet.
      For i = 2 To RowCount
      
          '3a) Increase volume for current ticker
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

          '3b) Check if the current row is the first row with the selected tickerIndex.
          'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then  
              tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          'End If
          End If

          '3c) check if the current row is the last row with the selected ticker
           'If the next row of ticker doesn't match, increase the tickerIndex.
          'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

              '3d Increase the tickerIndex.
              tickerIndex = tickerIndex + 1
          'End If
          End If
      Next i

      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
          Worksheets("All Stocks Analysis").Activate
          tickerIndex = i
    
          Cells(4 + i, 1).Value = tickers(tickerIndex)
          Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
          Cells(4 + i, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
      Next i
    
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

    End Sub

![Code_Run_Time_2017_Refactored](https://user-images.githubusercontent.com/106395288/173252906-a8621957-b8f6-49b3-ac7b-1cf5867514e2.png)
![Code_Run_Time_2018_Refactored](https://user-images.githubusercontent.com/106395288/173252661-6965f1fd-ae84-4236-a89d-168668205d5b.png)

### Stock Performance in 2017 and 2018
The stock (DQ), picked by Steve's parents, had a good performance in 2017. However, in 2018, DQ's return rate tumbled from 199.4% to -62.6%. DQ would not be a good investment option for steve's parents. 

For all green energy stocks' performances in 2017 and 2018, the green energy stocks had a better performance in 2017. In 2018, only two stocks had positive returns: EHPN and RUN. ENPH outperformed most of the other green energy stocks in both years. Even though its return rate dropped from 129.5% to 81.9% in 2018, as the general performance of the green energy stocks in 2018 was worse than in 2017, EHPN will be a good choice for investment. In 2017, RUN had an average performance among the other green stocks. However, in 2018 RUN is leading the way in both trading volume and return rate. Both ENPH and RUN are good options for Steve's parents to invest in.   

![Stock_Performance_2017](https://user-images.githubusercontent.com/106395288/173252682-934401be-f336-4a81-9c65-6ca5ac7f62bd.png)  
![Stock_Performance_2018](https://user-images.githubusercontent.com/106395288/173252693-413ed1a5-efe1-4a2f-b014-be095d733895.png)


## Summary 
### Advantages and Disadvantages of Refactoring Code
Refactoring code is the process of restructuring existing code while not changing its functionality.
>Refactoring is intended to improve the design, structure, and/or implementation of the software, while preserving its functionality. (Wikipedia)

The advantages of code refactoring are: 
- Improve the internal structure of the existing code
- Enhance code performance; reduce running time of the existing code
- Enhance readibility of the code

The disadvantages of code refactoring are:
- Affect the functionality of the original code
- Introduce new bugs into the code

### Apply Code Refactoring to the Original VBA Script
I encountered several errors as I had to refactor the code with different logic. But in the process of debugging, I got a better understanding of how VBA codes run. After I completed code refactoring, my code's performance enhanced a lot. The running time is five times faster than the original code now.
