Attribute VB_Name = "Module1"
' Module1 to summarize stocks per Quarter per sheet


Sub Stock_per_Quarter_per_sheet(ws As Worksheet)
    ' Create a variable to hold the counter
    Dim i As Long

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String

    ' Set an initial variable for holding the total stock volume per Ticker
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
  
    'Set an initial variable for holding open price, close price, Quarterly Change and Percent change for each Ticker
  
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Quarterly_Change As Double
    Quarterly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
  
    'Add the word "Ticker" to the Ticker Column Header
        ws.Cells(1, 9).Value = "Ticker"
  
    'Add the word "Quarterly Change" to the Quarterly Change Column Header
        ws.Cells(1, 10).Value = "Quarterly Change"
        
     'Add the word "Percent Change" to the Percent Change Column Header
        ws.Cells(1, 11).Value = "Percent Change"
        
     'Add the word "Total Stock Volume" to the Total Stock Volume Column Header
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        
  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  ' Set an initial variable for holding the Last Row of the sheet
  
    Dim LastRow As Long
    LastRow = 0
  ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Initialise Open price for the next ticker
    Open_Price = ws.Cells(2, 3).Value

  ' Loop through all Tickers
  For i = 2 To LastRow

    ' Check if we are still within the same credit card brand, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value
      
      'Set the close Price value
      Close_Price = ws.Cells(i, 6).Value
      
      ' Set the Quarterly_Change Value
      Quarterly_Change = Close_Price - Open_Price
      
      'Set the Percent Change Value
      
      Percent_Change = (Quarterly_Change / Open_Price)

      ' Add to the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ' Print the Quarterly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
      
       ' Print the Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change

      ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Reset Open Stock Value
       Open_Price = ws.Cells(i + 1, 3).Value
       
      ' Reset the Brand Total
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else
      ' Add to the Brand Total
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
    '-------------------------------------------------------------------------------------------
  
    ' Conditional Formatting of the summary table
    
    '--------------------------------------------------------------------------------------------
    ' Store Size of summary table as Summary_table_Row_max
    Dim Summary_table_Row_max As Long
    Summary_table_Row_max = Summary_Table_Row
    
    
    
    ' Conditional formatting is applied correctly and appropriately to the percent change column
    ' setting number to 2 decimal place
    ' example code { Range("A1:A10").NumberFormat = "0.00"}
    ' example code { Range("A1:A10").NumberFormat = "0.00%"}
    
    ws.Range("J2:J" & Summary_table_Row_max).NumberFormat = "0.00"
    ws.Range("K2:K" & Summary_table_Row_max).NumberFormat = "0.00%"
    ws.Range("L2:L" & Summary_table_Row_max).NumberFormat = "0"
    
    '---------------------------------------------------------------------------------------------
    ' Conditional formatting is applied correctly and appropriately to the quarterly change column
    ' if value < 0 Set the Cell Colors to Red
    ' example code{   Range("A2:A5").Interior.ColorIndex = 3 }
    ' if value > 0 Set the Font Color to Green
    ' example code{  Range("B1").Interior.ColorIndex = 4 )
    
    ' Create a variable to hold the counter to loop through the summary table
    Dim j As Long
    
    For j = 2 To Summary_table_Row_max
        If (ws.Range("J" & j).Value < 0) Then
            ws.Range("J" & j).Interior.ColorIndex = 3 ' set cell color to red
        Else
            ws.Range("J" & j).Interior.ColorIndex = 4 ' set cell color to green
        End If
    Next j
   
    
     
   
    
    
    
  
  
    ' --------------------------------------------------------------------------------------------------------
    ' Create the 2nd summary Table consisting of Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    
    ' ------------------------------------------------------------------------------------------------------
    
    ' Creating the variables to store Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    
    Dim Greatest_Percentage_increase As Double
    Dim Greatest_Percentage_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    ' Create variables to store the Ticker index and initialise them
    Dim Greatest_Percentage_increase_index As Long
    Dim Greatest_Percentage_Decrease_index As Long
    Dim Greatest_Total_Volume_index As Long
    Greatest_Percentage_increase_index = 2
    Greatest_Percentage_Decrease_index = 2
    Greatest_Total_Volume_index = 2
    
    
   
    
    
    Greatest_Percentage_increase = ws.Range("K" & 2).Value
    Greatest_Percentage_Decrease = ws.Range("K" & 2).Value
    Greatest_Total_Volume = ws.Range("L" & 2).Value
    
    For j = 3 To Summary_table_Row_max
        If (Greatest_Percentage_increase < ws.Range("K" & j).Value) Then
            Greatest_Percentage_increase = ws.Range("K" & j).Value
            Greatest_Percentage_increase_index = j
        End If
        If (Greatest_Percentage_Decrease > ws.Range("K" & j).Value) Then
            Greatest_Percentage_Decrease = ws.Range("K" & j).Value
            Greatest_Percentage_Decrease_index = j
        End If
        If (Greatest_Total_Volume < ws.Range("L" & j).Value) Then
            Greatest_Total_Volume = ws.Range("L" & j).Value
            Greatest_Total_Volume_index = j
        End If
    Next j
    
    '-----------------------------------------------------------------------------------------------------
    ' Create the Second summary Table
    '----------------------------------------------------------------------------------------------------
    
    
    ' Add row label for Greatest Percentage increase
    ws.Cells(2, 15).Value = "Greatest % increase"
    
     ' Add row label for Greatest Percentage decrease
    ws.Cells(3, 15).Value = "Greatest % decrease"
    
    ' Add row label for Greatest Total Volume
    ws.Cells(4, 15).Value = "Greatest Total Volume"
  
    
    'Add the word "Ticker" to the Ticker Column Header
        ws.Cells(1, 16).Value = "Ticker"
    ' Fill in the respective 3 greatest value Tickers
    ws.Cells(2, 16).Value = ws.Range("I" & Greatest_Percentage_increase_index).Value
    ws.Cells(3, 16).Value = ws.Range("I" & Greatest_Percentage_Decrease_index).Value
    ws.Cells(4, 16).Value = ws.Range("I" & Greatest_Total_Volume_index).Value
  
    'Add the word "Quarterly Change" to the Quarterly Change Column Header
    ws.Cells(1, 17).Value = "Value"
    
    ' Fill in the respective 3 greatest values
    ws.Cells(2, 17).Value = Greatest_Percentage_increase
    ws.Cells(2, 17).NumberFormat = "0.00%" ' format number to 2 decimal place %
    ws.Cells(3, 17).Value = Greatest_Percentage_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%" ' format number to 2 decimal place %
    ws.Cells(4, 17).Value = Greatest_Total_Volume
    ws.Cells(4, 17).NumberFormat = "0.00E+00"
    

End Sub



