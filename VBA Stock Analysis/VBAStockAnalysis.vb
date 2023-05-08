Sub stockdata()

'Loop Through all Sheets
For Each ws In Worksheets
          
'Create variable to Hold File Name, Last Row
Dim WorksheetName As String
Dim lastrowA As Long
Dim lastrowI As Long
        
'Determine the Last Row A
lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'Grab the WorksheetName
WorksheetName = ws.Name
        
'Set initial variables
Dim Ticker_Symbol As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double
Dim Summary_Table_Row As Integer
        
'Variable for greatest increase, greatest decrease and greatest total volume
Dim Greatest_Incr As Double
Dim Greatest_Decr As Double
Dim Greatest_Vol As Double
        
'Initialize total volume
 Total_Volume = 0
        
'Create Column Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
        
'Set initial Summary Table Row
 Summary_Table_Row = 2
        
'Loop through all the rows of Ticker Symbol
    For i = 2 To lastrowA
        
'Set the Open price
 If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
 
 Open_Price = ws.Cells(i, 3).Value
 
 End If
            
'Check if we are still within the same ticker symbol, if not...
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
'Calculate Total Stock Volume for Ticker Symbol
             
 Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
'Set the Close Price
 Close_Price = ws.Cells(i, 6).Value
                
'Set the Ticker Symbol
 Ticker_Symbol = ws.Cells(i, 1).Value
                
'Calculate Yearly Change
 Yearly_Change = Close_Price - Open_Price
                
'Calculate Percent Change
 If Open_Price <> 0 Then
 Percent_Change = Yearly_Change / Open_Price
 
 Else
                    
 Percent_Change = 0
 
 End If
               
                
'Conditional Formatting to fill the color

If Yearly_Change < 0 Then

'Set the fill colors to red
 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
 
 Else
 
'Set the fill colors to green
 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
End If
                
'Print the results in the Summary Table
ws.Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
ws.Cells(Summary_Table_Row, 12).Value = Total_Volume
                
'Print the percent change as percent format
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                
'Increased summary table row by 1
Summary_Table_Row = Summary_Table_Row + 1
    
'Get the next open price
Open_Price = ws.Cells(i + 1, 3).Value
    
'Reset Total Volume
 Total_Volume = 0
    
Else
    
Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
End If
    
Next i

'Now find last row in column I
lastrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Prepare for summary report
Greatest_Vol = ws.Range("L2").Value
Greatest_Incr = ws.Range("K2").Value
Greatest_Decr = ws.Range("K2").Value

'Loop through each row in column I
  For i = 2 To lastrowI
    
'Check if current total volume is greater than the greatest total volume so far
If ws.Range("L" & i).Value > Greatest_Vol Then

'If so, update the greatest total volume and corresponding ticker symbol
Greatest_Vol = ws.Range("L" & i).Value
ws.Range("P4").Value = ws.Range("I" & i).Value

End If
    
'Check if current percentage increase is greater than the greatest percentage increase so far
If ws.Range("K" & i).Value > Greatest_Incr Then

'If so, update the greatest percentage increase and corresponding ticker symbol
Greatest_Incr = ws.Range("K" & i).Value
ws.Range("P2").Value = ws.Range("I" & i).Value
 
End If
    
'Check if current percentage decrease is smaller than the greatest percentage decrease so far
If ws.Range("K" & i).Value < Greatest_Decr Then

'If so, update the greatest percentage decrease and corresponding ticker symbol
 Greatest_Decr = ws.Range("K" & i).Value
 ws.Range("P3").Value = ws.Range("I" & i).Value
 
End If

Next i


'Write summary report results in worksheet
ws.Range("Q2").Value = Format(Greatest_Incr, "Percent")
ws.Range("Q3").Value = Format(Greatest_Decr, "Percent")
ws.Range("Q4").Value = Format(Greatest_Vol, "Scientific")


'Adjust column width automatically
ws.Columns("A:Z").AutoFit


Next ws


End Sub
