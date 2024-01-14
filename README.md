# VBA-challenge

# Instructions:

A) Create a VBA script that loops through every worksheet and outputs the following information (1 to 5):  
(1) The ticker symbol  
(2) Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year  
(3) The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year  
(4) The total stock volume of the stock  
(5) Add functionality to the script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"  

B) Use conditional formatting that will highlight positive change in green and negative change in red

C) Take screenshot of the result, and the two screenshots should match the screenshots provided by the Course instruction

# Main Concept Tested:

1) For Each and For Loops: 
The script uses a For Each loop to iterate through each worksheet (ws), and For loop to iterate through rows within each worksheet.
  *' For Each loop,
  *' Reviewed [documentation](https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook), and 03-VBA-Scripting class activity 07-Stu_Census_Pt1 to refresh my memory to loop through each worksheet.

2) Variables and Data Types:
Various variables like WorksheetName, Ticker_Symbol, Volume_Total, Open_Price, Close_Price, etc., are declared and used to store and manipulate data.
  *' I have tried to Dim Volume_Total As Long, but has resulted in error message "overflow". Therefore, I change the data type to double, as a 'Double' can handle very large numbers, including those that might cause an overflow error with 'Long'.

3) Conditional Statements:
Conditional statements (If...Then...Else) are used to check if the current row represents the end of a ticker symbol group.

# Final Screenshots:
![Assignment_Screenshot1](https://github.com/vanillatyy1/VBA-challenge/blob/main/Stock_Ticker_Screenshot_1.jpg)
![Assignment_Screenshot2](https://github.com/vanillatyy1/VBA-challenge/blob/main/Stock_Ticker_Screenshot_2.jpg)

Note that the arrow was not a requirement for the VBA script; it was included in response to the Class Instruction screenshot for alignment purposes.

# Some good-to-remember-for-the-future Technique Used:
i) LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Line of code to find the last non-empty row in a specific column of the worksheet in Excel VBA.

ii) Website for color guide

See website for color guides: http://dmcritchie.mvps.org/excel/colors.htm. 
  *'I have learnt about this link during the 3rd VBA class, activity 02-Ins_Formatter.
  *'Red
  Interior.ColorIndex = 3
  *'Green
  Interior.ColorIndex = 4, as in

    ' --------------------------------------------
    ' CONDITIONAL FORMATTING
    ' --------------------------------------------

    ' Apply conditional formatting to column K
    If ws.Cells(Summary_Table1_Row, 10).Value > 0 Then
    
    ' Set the Cell Colors to Green
    ws.Range("J" & Summary_Table1_Row).Interior.ColorIndex = 4
                
    Else
    
    ' Set the Cell Colors to Red
    ws.Range("J" & Summary_Table1_Row).Interior.ColorIndex = 3

iii) Greatest Total Volume Value Format
To match the image provided by Class instruction, no change of Format is required to display 1.69E+12 into 1689539560106.
If it became a requirement, then we would need to add the following to the script

    ' --------------------------------------------
    ' FORMAT CELLS AS NUMBER
    ' --------------------------------------------

     ' Format cells in column L as Number w/ no decimal places
     ws.Range("L:L").NumberFormat = "0"
        
     ' Format cell P4 a Number w/ no decimal places
     ws.Range("P4").NumberFormat = "0"


