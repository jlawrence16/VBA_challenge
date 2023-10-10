Option Explicit

Sub vba_challenge():

Dim ws As Worksheet

For Each ws In Worksheets


Dim open_price As Double
Dim close_price As Double
Dim i As Double
Dim stock_volume__total As Double
Dim ticker_symbol_name As String
Dim counter As Integer
Dim lrow As Double
Dim yearly_change As Double
Dim percentage_change As Double




lrow = Range("A" & Rows.Count).End(xlUp).Row


    'New column headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

  ' Set an initial variable for holding the total stock volume per ticker symbol
  
        stock_volume__total = 0

  ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
    
    'Set opening price for first iteration
        open_price = Cells(2, 3).Value


  ' Loop through all ticker symbol data
        For i = 2 To lrow
   

            ' Check if we are still within the same ticker symbol, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' Set the Ticker Symbol
                ticker_symbol_name = ws.Cells(i, 1).Value
      
                ' Add to the Stock value Total
                stock_volume__total = stock_volume__total + ws.Cells(i, 7).Value
                                      
                'Get close price
                close_price = ws.Cells(i, 6).Value
                
                yearly_change = close_price - open_price
                
                percentage_change = (close_price - open_price) / open_price
                
                
                ' Print the yearly change Total to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = yearly_change

                 ' Print the Ticker Symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker_symbol_name
                
                ' Print the percentage change to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = percentage_change

                 ' Print the Stock Volume Total to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = stock_volume__total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the stock value total
                stock_volume__total = 0
                
                'Shift down one cell to get the opening price for next iteration
                open_price = ws.Cells(i + 1, 3).Value
                   
            ' If the cell immediately following a row is the same brand...
            Else
               
            ' Add to the Stock value Total
            stock_volume__total = stock_volume__total + ws.Cells(i, 7).Value


        End If

    Next i
    
   'Format and Conditional Format new columns
   
        'Autowidth Columns
        ws.Columns("I:L").EntireColumn.AutoFit
        
        'Format percent change column as percentage
        ws.Columns("K").NumberFormat = "0.00%"
        
        'Loop to colour cells green for yearly increase and red for yearly decrease
        Dim j As Double
        Dim lrow_summary_table As Integer
        lrow_summary_table = ws.Range("J" & Rows.Count).End(xlUp).Row
        
        
            For j = 2 To lrow_summary_table
            
            'If Then statement to colour green or red as appropriate
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            
            Next j
    
    
        'Calcluate and print greatest % increase, greatest % decrease and greatest total volume
        
        'Print new columns titles
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
        'Print new row titles
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
        'Autofit column width
            ws.Columns("O").EntireColumn.AutoFit
            
        'Calcluate greatest % increase, greatest % decrease and greatest total volume
            
            Dim greatest_per_increase As Double
            Dim greatest_per_decrease As Double
            Dim greatest_volume_increase As Double
            
            greatest_per_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & lrow_summary_table))
            
            
            greatest_per_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lrow_summary_table))
            
            greatest_volume_increase = Application.WorksheetFunction.Max(ws.Range("L2:L" & lrow_summary_table))
            
        'Print results to table, autofit column width and format appropraite cells as percentage
            
            ws.Cells(2, 17).Value = greatest_per_increase
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ws.Cells(3, 17).Value = greatest_per_decrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
           ws.Cells(4, 17).Value = greatest_volume_increase
           ws.Columns("Q").EntireColumn.AutoFit
            
        'Identify and print associated ticker symbols
            
            Dim k As Integer
            Dim ticker_greatest_per_increase As String
            Dim ticker_greatest_per_decrease As String
            Dim ticker_greatest_volume_increase As String
            
            
                For k = 2 To lrow_summary_table
                    
                    If ws.Cells(k, 11) = greatest_per_increase Then
                        ticker_greatest_per_increase = ws.Cells(k, 11).Offset(, -2)
                        ws.Cells(2, 16).Value = ticker_greatest_per_increase
                    
                    ElseIf ws.Cells(k, 11) = greatest_per_decrease Then
                        ticker_greatest_per_decrease = ws.Cells(k, 11).Offset(, -2)
                        ws.Cells(3, 16).Value = ticker_greatest_per_decrease
                    
                    ElseIf ws.Cells(k, 12) = greatest_volume_increase Then
                        ticker_greatest_volume_increase = ws.Cells(k, 12).Offset(, -3)
                        ws.Cells(4, 16).Value = ticker_greatest_volume_increase
                    
                    End If
                    
                Next k
    Next ws
                                     
End Sub




