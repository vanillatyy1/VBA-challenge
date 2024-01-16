Attribute VB_Name = "Module1"
Sub stock_ticker()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------

For Each ws In Worksheets

        ' Create a Variable to Hold File Name, Last Row, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        'MsgBox WorksheetName
        
        ' Print headers to I1:L1, O1:P1, N2:N4
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Year Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
  ' Set an initial variable for holding the ticker
  Dim Ticker_Symbol As String
    
  ' Set an initial variable for holding the volume total per ticker
  Dim Volume_Total As Double
  Volume_Total = 0
        
  ' Set an initial variable for holding open price
  Dim Open_Price As Double
  Open_Price = 0

  ' Set an initial variable for holding the close price
  Dim Close_Price As Double
  Close_Price = 0
  
    ' Keep track of the location for each ticker in Summary Table1
    ' Summary Table = column I:L
    
  Dim Summary_Table1_Row As Integer
  Summary_Table1_Row = 2

    ' Date start
    Dim DateStartRow As Long
    DateStartRow = 2
    
    'Set initial variable for holding greatest increase, greatest decrease, greatest increase ticker, greatest decrease ticker,
    'greatest total volume, the ticker symbol associated with the greatest total volume value
    
    Dim Greatest_increase As Double
    Dim Greatest_decrease As Double
    Dim Greatest_volume As Double
    Dim Greatest_increase_ticker As String
    Dim Greatest_decrease_ticker As String
    Dim Greatest_volume_ticker As String
    
    Greatest_increase = 0
    Greatest_decrease = 0
    Greatest_volume = 0
  
    ' Loop through all rows
  For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
    ' Set the Ticker symbol
    Ticker_Symbol = ws.Cells(i, 1).Value 'During the 1st loop, at this point, VBA is already looking at cell A252, which is the last AAB before the ticker changes to AAF, so i = 252
    
    ' Find the Opening price & Closing Price
    Open_Price = ws.Cells(DateStartRow, 3).Value
    Close_Price = ws.Cells(i, 6).Value
    
    ' Basic math to find the value for Year Change and Percent change
    Year_Change = Close_Price - Open_Price
    Percent_change = Year_Change / Open_Price
    
    ' Add the last volume value into the Total
    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
       
    ' Print the Ticker Symbol in Summary Table1
    ws.Range("I" & Summary_Table1_Row).Value = Ticker_Symbol
    
    ' Print the Year Change
    ws.Range("J" & Summary_Table1_Row).Value = Year_Change
    
    ' Print the Percentage change
    ws.Range("K" & Summary_Table1_Row).Value = Percent_change
    
    ' Print the Total Stock Volume to Summary Table1
    ws.Range("L" & Summary_Table1_Row).Value = Volume_Total

    If Percent_change > Greatest_increase Then
    Greatest_increase = Percent_change 'The Percent_change will become the Greatest_increase, and Greatest_increase will not be 0 anymore
    Greatest_increase_ticker = Ticker_Symbol 'Greatest_increase_ticker = Cell O2
            
    ElseIf Percent_change < Greatest_decrease Then 'In the first loop, if Percent_change is < the current Greatest decrease, which is 0, Then
    Greatest_decrease = Percent_change 'then the Percent_change will become the new Greatest_decrease
    Greatest_decrease_ticker = Ticker_Symbol 'Greatest_decrease_ticker = Cell O3
    
    End If
            
    If Volume_Total > Greatest_volume Then 'If Total_volume, for instance, in 1st loop, 765628638 > the current value in Greatest_volume which is 0, Then
    Greatest_volume = Volume_Total 'the Greatest_volume, which is cell P4, will be replaced by the Total_volume, which in the 1st loop, will be 765628638
    Greatest_volume_ticker = Ticker_Symbol
    
    End If
    
    ' --------------------------------------------
    ' CONDITIONAL FORMATTING
    ' --------------------------------------------

    ' Apply conditional formatting to column J (Year Change)
    If ws.Cells(Summary_Table1_Row, 10).Value > 0 Then
    
    ' Set the Cell Colors to Green
    ws.Range("J" & Summary_Table1_Row).Interior.ColorIndex = 4
                
    Else
    
    ' Set the Cell Colors to Red
    ws.Range("J" & Summary_Table1_Row).Interior.ColorIndex = 3
    
    End If
    
    ' --------------------------------------------
    ' PLAN FOR NEXT TICKET SYMBOL
    ' --------------------------------------------
    
    ' Add one to the Summary Table1 row
    Summary_Table1_Row = Summary_Table1_Row + 1 'in the first loop, this will be 252 + 1 = 253. In the next loop, the VBA will start looking at row 253
          
    ' Reset the Volume Total
    Volume_Total = 0
    
    DateStartRow = i + 1
    
    ' If the cell immediately following a row is the same ticker symbol...
    Else

    ' Add to Volume Total
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
    End If

  Next i
  
    ' --------------------------------------------
    ' PREPARE AND PRINT DATA FOR SUMMARY TABLE 2
    ' --------------------------------------------
  
    ' Print value in Summary_Table2 (O2:P4)
    ws.Range("P2").Value = Greatest_increase_ticker
    ws.Range("P3").Value = Greatest_decrease_ticker
    ws.Range("P4").Value = Greatest_volume_ticker
    
    ws.Range("Q2").Value = Greatest_increase
    ws.Range("Q3").Value = Greatest_decrease
    ws.Range("Q4").Value = Greatest_volume
    
    ' --------------------------------------------
    ' FORMAT CELLS AS PERCENTAGE
    ' --------------------------------------------
    
    ' Format Percentage for Column K (Percentage Change)
    ws.Range("K:K").NumberFormat = "0.00%"
      
    ' Format Percentage for Cell Q2 (Greatest % Increase)
    ws.Range("Q2").NumberFormat = "0.00%"
        
    ' Format Percentage for Cell Q3 (Greatest % Decrease)
    ws.Range("Q3").NumberFormat = "0.00%"
       

Next ws


End Sub


