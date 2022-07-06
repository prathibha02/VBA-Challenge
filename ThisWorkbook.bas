VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabetical_testing()


Dim ws As Worksheet


'Define the varibles

Dim summary_row As Long
Dim ticker As String
Dim opening_price As Currency
Dim closing_price As Currency
Dim stock_volume As LongLong
Dim yearly_change As Double
Dim percent_change As Double
Dim column_heading As String
Dim row_heading As String
Dim current_year As Integer
Dim min_max_results As Integer
Dim lrow As Integer
Dim current_stock_volume As LongLong
Dim l_results_row As Integer
Dim work_sheet_count As Integer

'Initialize variables

summary_row = 2
opening_price = Range("C2").Value
closing_price = 0
stock_volume = 0
yearly_change = 0
percent_change = 0
current_stock_volume = 0
l_results_row = 0
current_year = 0
min_max_results = 2
'Set ws = Worksheets("A")

For work_sheet_count = 1 To 6

'Look into the current worksheet

ActiveWorkbook.Worksheets(work_sheet_count).Activate

'Message out the sheet name
MsgBox ActiveWorkbook.Worksheets(work_sheet_count).Name

'Move the current worksheet to a varible ws
Set ws = Worksheets(ActiveWorkbook.Worksheets(work_sheet_count).Name)

'Get the last row count of the data

lrow = Cells(Rows.Count, 1).End(xlUp).Row

'Write out the column labels

 column_heading = "Ticker"
 Cells(1, 10) = column_heading
 Cells(1, 17) = column_heading
 
 column_heading = "Yearly Change"
 Cells(1, 11) = column_heading
  
 column_heading = "Percent Change"
 Cells(1, 12) = column_heading
  
 column_heading = "Total Stock Volume"
 Cells(1, 13) = column_heading
  
 column_heading = "Value"
 Cells(1, 18) = column_heading
  
 'Write out the Row explanation for summaries
 
  row_heading = "Greatest % Increase"
  Cells(2, 16) = row_heading
  
  row_heading = "Greatest % Decrease"
  Cells(3, 16) = row_heading
  
  row_heading = "Greatest Total volume"
  Cells(4, 16) = row_heading

'Loop through first to last row
 
  For i = 2 To lrow
    
    'Get the current year
     current_year = Left(Cells(i, 2).Value, 4)

     'Loop through the data for the current year
        
          If current_year = Left(Cells(i + 1, 2).Value, 4) Then

       ' If the ticker symbol is the same
       
          If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
             
             'Add the stock volume
              current_stock_volume = Cells(i, 7).Value
  
               stock_volume = stock_volume + current_stock_volume
   
               current_stock_volume = 0
   
           Else
  
           'Otherwise (if the ticker symbol is not the same, move the ticker symbol to the variable ticker
               ticker = Cells(i, 1)
     
            ' Add to the Stock Volume
    
                current_stock_volume = Cells(i, 7).Value
                stock_volume = stock_volume + current_stock_volume

             ' Print the Ticker in the Summary
                Range("J" & summary_row).Value = ticker
              
             '  Move the closing price to the varible
                closing_price = Cells(i, 6).Value
        
             ' Calculate the yearly change, percent change
              
                yearly_change = closing_price - opening_price
                percent_change = (yearly_change / opening_price)
        
        
             ' Print the Yearly Change to the Summary
               
               'If the change is positive, make the background of the cell green
               
                If yearly_change >= 0 Then
        
                  Range("K" & summary_row).Interior.ColorIndex = 4
                  Range("K" & summary_row).Value = yearly_change
         
               'Or else make the background of the cell red
               
                 Else
        
                  Range("K" & summary_row).Interior.ColorIndex = 3
                  Range("K" & summary_row).Value = yearly_change
        
                 End If
         
        
           'Print the Yearly Change to the Summary
       
               Range("L" & summary_row).Value = FormatPercent(percent_change, [2])
        
        
            'Print the Stock Volume to the Summary
            
               Range("M" & summary_row).Value = stock_volume

            ' Add one to the summary row
            
               summary_row = summary_row + 1
      
             ' Reset the varibles to be used again
             
                stock_volume = 0
                opening_price = 0
                closing_price = 0
                percent_change = 0
                yearly_change = 0
                current_stock_volume = 0
                
        
             ' Add the opening_price of the next ticker
             
               opening_price = Cells(i + 1, 3).Value
         
        ' End of the current ticker price
          End If
        
           current_year = 0
           
       ' Before the next year
           Else
     
            summery_row = summery_row + 1
            Cells(2, 9) = current_year
            Cells(2, 15) = current_year
            'summery_row = summery_row + 1
            current_year = 0
    End If
           
    
   Next i
     
     
     summary_row = 2
      i = 2
    'Calculate the number of rows in the summary table
    
     
    
     'Loop through the Aggregates for each ticker
      For i = 2 To 200
      'return highest number in a range
       ws.Range("R2") = FormatPercent(Application.WorksheetFunction.Max(ws.Range("L2:L100")), [2])

        
         'return highest number in a range
         ws.Range("R3") = FormatPercent(Application.WorksheetFunction.Min(ws.Range("L2:L100")), [2])

          'return highest number in a range
          ws.Range("R4") = Application.WorksheetFunction.Max(ws.Range("M2:M100"))
     
     Next i
         
          l_results_row = Cells(Rows.Count, 10).End(xlUp).Row
          
          For i = 2 To l_results_row
         
          If Cells(min_max_results, 18).Value = Cells(i, 12).Value Then
            Cells(min_max_results, 17) = Cells(i, 10)
            min_max_results = min_max_results + 1
                     
                   
          
         ElseIf Cells(min_max_results, 18).Value = Cells(i, 12).Value Then
            Cells(min_max_results, 17) = Cells(i, 10)
             min_max_results = min_max_results + 1
         
       
          
          
           ElseIf Cells(min_max_results, 18).Value = Cells(i, 13).Value Then
             
             Cells(min_max_results, 17) = Cells(i, 10)
              min_max_results = min_max_results + 1
         
           End If
       
       
           
     Next i
      'min_max_results = 0
     ' l_results_row = 0
     ' l_row = 0
       
  Next work_sheet_count

 
    
          
       
 
       
          

End Sub
