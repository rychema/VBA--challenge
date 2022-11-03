Sub Multiple_year_stock_data()

'Loop trough all sheets

 For Each WS In Worksheets
 WS.Activate
 
 'Set new titles for columns and rows
 
 WS.Cells(1, 9).Value = "ticker"
 WS.Cells(1, 10).Value = "yearly chnage"
 WS.Cells(1, 11).Value = "percent change"
 WS.Cells(1, 12).Value = "total stock volume"
 WS.Cells(1, 16).Value = "ticker"
 WS.Cells(1, 17).Value = "value"
 WS.Cells(2, 15).Value = "Greates% Increase"
 WS.Cells(3, 15).Value = "Greatest% Decrease"
 WS.Cells(4, 15).Value = "Greates Total Volume"
 
'Define Variables

 
Dim ticker_symbol As String
Dim ticker_total As Double
ticker_total = 0


Dim yearly_change As Double
Dim percent_change As Double
Dim open_price As Double
Dim close_price As Double


Dim Greates_Inc As Double
Dim Greates_Dec As Double
Dim Greates_Vol As Double



'Set open and close price

open_price = WS.Cells(2, 3).Value

close_price = WS.Cells(2, 6).Value
 



'Define summary table

Dim summary_table_row As Integer

summary_table_row = 2


'Define last row

Dim lastrow As Double

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

   MsgBox ("LastRow in column 1 is " & lastrow)
   
   
'Loop through all ticker

 For i = 2 To lastrow
 
' Check if we are still within the same ticker

    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
ticker_total = ticker_total + Cells(i, 7).Value
        
    If (open_price = 0) Then
    
open_price = Cells(i + 1, 3).Value

    End If
    
Else

'If we not within same ticker...

ticker_symbol = Cells(i, 1).Value

Range("I" & summary_table_row).Value = ticker_symbol


ticker_total = ticker_total + Cells(i, 7).Value

Range("L" & summary_table_row).Value = ticker_total

close_price = WS.Cells(i, 6).Value

yearly_change = close_price - open_price

Range("J" & summary_table_row).Value = yearly_change

  If (open_price = 0) Then
  
Cells(summary_table_row, 11).Value = 0
        
'If open year is more than 0

Else

    percent_change = (yearly_change / open_price) * 100
    
    
    Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")

End If

 'Background  for yearly_change, If we are on a cell that is positive then color is green
 
    If (yearly_change > 0) Then
    
WS.Range("J" & summary_table_row).Interior.ColorIndex = 4

' Otherwise color is red

    ElseIf (yearly_change <= 0) Then
    
WS.Range("J" & summary_table_row).Interior.ColorIndex = 3

    End If

Cells(summary_table_row, 12).Value = ticker_total

'Add to the summary table

        summary_table_row = summary_table_row + 1
        
'Reset Ticker Total

    ticker_total = 0
        
    open_price = Cells(i + 1, 3)
            
    End If
        
        

    Next i
        

'Prints the biggest value in perecent change ( column K)

        Greates_Inc = WorksheetFunction.Max(Range("K:K"))
         Range("Q2").Value = Greates_Inc
         Range("Q2").NumberFormat = "0.00%"

        
'Prints the smallest value in perecent change ( column K)

       Greates_Dec = WorksheetFunction.Min(Range("K:K"))
        Range("Q3").Value = Greates_Dec
        Range("Q2").NumberFormat = "0.00%"
        
'Prints the biggest value in total stock (Column L)

        Greates_Vol = WorksheetFunction.Max(Range("L:L"))
        Range("Q4").Value = Greates_Vol
        Range("Q4").NumberFormat = "0.00E+0"
        
'Prints ticker symbol into ticker column (column P)


        Range("P2") = "=Index(I:I,match(Q2,K:K, 0))"
        Range("P3") = "=Index(I:I,match(Q3,K:K, 0))"
        Range("P4") = "=Index(I:I,match(Q4,L:L, 0))"

'Autofit the column for all ranges

        Range("I:Q").EntireColumn.AutoFit
        Next WS
End Sub



