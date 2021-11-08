# VBA_Challenge
## Overview of Project: Explain the purpose of this analysis:
 We'll update, or remodel, the Stock Market Dataset with VBA solution code to loop over all of the data once in order to gather a whole dataser in this project and analysis. Then we'll see if restructuring your code resulted in the VBA script running quicker. Finally, we want to make the code more efficient—by reducing the number of steps, using less memory, or enhancing the logic of the code to make it easier to understand for future users.
## Results:
 1. The tickerIndex is set equal to zero before looping over the rows.
   - Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four     
    different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.
     tickerIndex = 0
 2. Arrays are created fortickerVolumes, tickerStartingPrices, and tickerEndingPrices.
     Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
 3. The tickerIndex is used to access the stock ticker index for the , tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

     For i = 2 To RowCount
        
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
     If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
     End If
     If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
              tickerIndex = tickerIndex + 1
      End If
    
      Next i
    4. Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
      Worksheets("All Stocks Analysis").Activate
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
      Next i
      
    5. Formatting
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
   ### Picture of the stock performance between 2017 and 2018:
  ![image](https://user-images.githubusercontent.com/93515126/140689135-d72272f6-8610-43dc-a13e-07e818acb2f3.png)
  ![image](https://user-images.githubusercontent.com/93515126/140689557-ac5c6f51-c104-4b39-b144-54997385ab17.png)

  ## Summary: 
   1. What are the advantages or disadvantages of refactoring code?
    The main benefit of refactoring code is that it makes it more efficient. The main downside of restructuring code is that you are potentially making code that is currently     useable useless if you do not restructure it appropriately. As a result, it's always a good idea to keep your original code just in case you can't rework it.
   2. How do these pros and cons apply to refactoring the original VBA script?
    The main benefit of restructuring code in a VBA script is that you may keep as much of the original code as you like and utilize various modules to place your new code       alongside your old code. The main drawback of restructuring code in VBA script is that you will struggle to reorganize your code if you do not have a thorough grasp of       the syntax, as syntax matters so much more when attempting to make your code more efficient.
