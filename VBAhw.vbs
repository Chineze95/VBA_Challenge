Attribute VB_Name = "Module1"
Sub StockMarket()

'Make sure to set the worksheet
'define each data

 Dim ws As Worksheet
 Dim ticker As String
 Dim yearly_change As Double
 Dim percent_change As Double
 Dim total_stock_volume As Double
 Dim openprice As Double
 Dim closeprice As Double
 Dim lastrow As Long
 Dim lastcol As Integer
 Dim i As Long
 
'loop over each worksheet in the workbook
 'For Each ws In Worksheets
 
 'Count the number of rows
 'lastrow = Cells(Rows.Count, "A").End(x1Up).Row
 lastrow = Cells(Rows.Count, "A").End(xlUp).Row
 
 'Count the number of columns
'lastcol = ws.Cells(Columns.Count, 1).End(x1ToLeft).Column
 
'Create the heading using the range or cell function
 Range("I1").Value = "Ticker"
 Range("J1").Value = "Yearly Change"
 Range("K1").Value = "Percent Change"
 Range("L1").Value = "Total Stock Volume"
 
'Initialize variables for each worksheet.
    ticker = ""
    yearly_change = 0
    percent_change = 0
    total_stock_volume = 0
    openprice = 0
    closeprice = 0
    
    
'Set the location for the following variables
 Dim summarytablerow As Long
 summarytablerow = 2
 
'Set up values for for the stock
 openprice = Cells(2, 3).Value
 
 
'Declare a variable code last row that will store the number of rows
'Start looping
'F statement listed in a for loop
'Make sure the ticker name is still the same
 For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    total_stock_volume = 0
    

 
'Create the ticker name
         tickername = Cells(i, 1).Value
 
'Get the end of the year closing price for ticker
        closeprice = Cells(i, 6).Value
            
'Get yearly change value
        yearly_change = closeprice - openprice
        
       Else 'Calculate next open price
 total_stock_volume = total_stock_volume + Cells(i, 7).Value
 If openprice = 0 Then
 openprice = Cells(i, 3).Value
 End If
 
        
        End If

'Make sure the open price does not come out to zero
 If openprice <> 0 Then
 yearly_change = (yearly_change / openprice) * 100
 
 
 End If
 
        
        
'Run this if we get to a different ticker in the list.
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
'Print ticker name for column I
 Range("I" & summarytablerow).Value = tickername
 
'Print yearly change for column J
 Range("J" & summarytablerow).Value = yearly_change
 
'Print percent change for column K
 Range("K" & summarytablerow).Value = (CStr(percent_change) & "%")
 
'Print total stock volume column L
 Range("L" & summarytablerow).Value = total_stock_volume
 
 
 End If
            
'Fill in the color for yearly change and specify what the different colors are for
 If yearly_change > 0 Then
 Range("J" & summarytablerow).Interior.ColorIndex = 4
 
 ElseIf yearly_change <= 0 Then
 Range("J" & summarytablerow).Interior.ColorIndex = 3
 
 End If
 
'Add 1 to summary table row count
 summarytablerow = summarytablerow + 1
 
 
   
  Next i
 
  'Next ws
  
End Sub



