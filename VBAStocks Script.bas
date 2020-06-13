Attribute VB_Name = "Module1"
Sub Stocks()

For Each ws In Worksheets


Dim LastRow As Long
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row

Dim TickerName As String

'set an initial variable for tracking row in table
Dim RowCounter As Integer
RowCounter = 2

'set initial variable for holding the value of the stock
Dim TotalStock As Double
TotalStock = 0

'set initial variable for holding the value of the yearly change
Dim YearlyOpen As Double
YearlyOpen = ws.Cells(2, 3).Value
Dim YearlyClose As Double
Dim FirstOpen As Double

Dim PercentChange As Double

Dim GreatestIncrease As Double
'GreatestIncrease = 0
Dim GreatestDecrease As Double
'GreatestIncrease = 0

Dim Max_Percentage_Increase As Double
Dim Max_Percentage_Decrease As Double
Dim Greatest_Total_Volume As Double

Dim Max_Increase_Row As Double
Dim Max_Decrease_Row As Double
Dim Max_Volume_Row As Double




'add headers to columns
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'set column width to fit for columns L, O, P, and Q
ws.Columns("L:Q").EntireColumn.AutoFit








For I = 2 To LastRow
    
    ' check if we are still within the same ticker, if not:
    If ws.Cells(I, 1) <> ws.Cells(I + 1, 1) Then
    
        'set the ticker name
        TickerName = ws.Cells(I, 1)
        'print the ticker name in the summary table
        ws.Range("I" & RowCounter) = ws.Cells(I, 1)
        
        'add to the stock total
        TotalStock = TotalStock + ws.Cells(I, 7)
        'print the stock total in summary table
        ws.Range("L" & RowCounter) = TotalStock


               
        'reset TotalStock total
        TotalStock = 0
     
            
        '-----------------------------------
        ' YEARLY CHANGE & PERCENT CHANGE CODE
        '-----------------------------------
             
        
        'assign yearlyclose a value
        YearlyClose = ws.Cells(I, 6).Value
        
        'print yearly change in summary table
        ws.Range("J" & RowCounter) = YearlyClose - YearlyOpen
            
        'percent change code
        If YearlyOpen = 0 Then

            PercentChange = YearlyClose - YearlyOpen
            
        Else
        
            PercentChange = (YearlyClose - YearlyOpen) / YearlyOpen
        
        End If
        
        ws.Range("K" & RowCounter) = Format(PercentChange, "Percent")
        

        
          
        'reset yearly open to original value
        YearlyOpen = ws.Cells(I + 1, 3)
         
        'add one to the table row
        RowCounter = RowCounter + 1
                
    'if the cell immediately following a row is the same ticker..
    Else
    
        'add to the total stock
        TotalStock = TotalStock + ws.Cells(I, 7).Value
        
    End If
    
Next I

       '-----------------------
       'Determine Greatest Percentage Increase/Decrease, Greatest Total Volume and populate tickers
       '-----------------------
        
       'determine greatest percentage increase and print
       Max_Percentage_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
       ws.Range("Q2") = Max_Percentage_Increase
       ws.Range("Q2").NumberFormat = "0.00%"
              
              
       'determine greatest percentage decrease and print
       Max_Percentage_Decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
       ws.Range("Q3") = Max_Percentage_Decrease
       ws.Range("Q3").NumberFormat = "0.00%"
        
       'calculate total volume and print
       Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
       ws.Range("Q4") = Greatest_Total_Volume
       ws.Range("Q4").NumberFormat = "0"
       
       'populate tickers
       'get row for greatest percentage change
       Max_Increase_Row = Application.WorksheetFunction.Match(Max_Percentage_Increase, ws.Range("K:K"), 0)
       'then set ticker to the row, column and print
       ws.Range("P2").Value = ws.Cells(Max_Increase_Row, 9).Value
       
       'get row for greatest decrease change
       Max_Decrease_Row = Application.WorksheetFunction.Match(Max_Percentage_Decrease, ws.Range("K:K"), 0)
       'then set ticker to the row, column and print
       ws.Range("P3").Value = ws.Cells(Max_Decrease_Row, 9).Value
       
       'get row for greatest total volume
       Max_Volume_Row = Application.WorksheetFunction.Match(Greatest_Total_Volume, ws.Range("L:L"), 0)
       'then set ticker to the row, column and print
       ws.Range("P4").Value = ws.Cells(Max_Volume_Row, 9).Value
       



'-----------------------------------------
'conditional formatting for yearly change
'-----------------------------------------

For I = 2 To RowCounter - 1
    
    If ws.Cells(I, 10) <= 0 Then
        
        ws.Cells(I, 10).Interior.ColorIndex = 3
        
    Else
        
        ws.Cells(I, 10).Interior.ColorIndex = 4
    
    End If
    
Next I


Next ws

MsgBox ("Run Complete")


End Sub
