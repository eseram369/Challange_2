# Challange_2
## stock_analysis with Excel (VBA)
### Overview of Project
#### The purpose
##### The purpose is to editor the original code to make it more efficient. The goal of using Excel VBA is to help Steve analyze an entire dataset and collect certain stock information between 2017 and 2018 and determine whether the stock is worth investing in or not? 
#### Background 
The data gives the two tables for 2017 and 2018 with stock information on 12 different stocks. The stock information includes a ticker value, the date the stock has an issue, the open, closing, adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock
#### Result
##### Analysis

    1a) Create a ticker Index
     Dim tickerIndex As Long
    
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
     For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, "H").Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, "A").Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, "F").Value
        End If
                
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, "A").Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, "F").Value
        
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        
        Cells(4 + i, "A").Value = tickers(tickerIndex)
        Cells(4 + i, "B").Value = tickerVolumes(tickerIndex)
        Cells(4 + i, "C").Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
    
    
  ####  Result 
      
    
   ![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/90746609/135791201-e850b080-9d26-4e86-b654-7d3e352f0bb2.jpg)


![VBA_Challenge_2018png](https://user-images.githubusercontent.com/90746609/135791178-9aa49509-e183-4d23-b68a-bae9c6f2c4b4.jpg)

#### Summery
1.What are the advantages or disadvantages of refactoring code?
The advantages of refactoring make the code more organized.  Also, it was faster,  debugging smooth. It is easier to read for other users. The disadvantages take a long time to learn and can be risky.
2.How do these pros and cons apply to refactoring the original VBA script?
The result can be efficient, clear, and fast to run. 

