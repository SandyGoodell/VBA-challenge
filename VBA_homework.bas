Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.

'The ticker symbol. --this is a string variable and will be 1st column, this is the what we are tracking
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'yearly_chg is the diff between first_openpr - last_clospr for each ticker, use first_openpr for the denomitator and
    'yearly_chg as the numerator.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'yearly_chg / first_openpr - this needs to be checked for 0 value for error trapping
'The total stock volume of the stock.
    '
'You should also have conditional formatting that will highlight positive change in green and negative change in red.




Sub Loop_Output()

    For Each ws In Worksheets
    
        'Print headers on worksheet
        ws.Cells(1, 10).Value = "Tickers"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percentage of Yearly Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        
    
        'Set an intial variable for ticker symbol,opening price, closing price, yearly change,% chg per ticker symbol, total volume per ticker
        Dim ticker As String
        Dim first_openpr As Double
        Dim last_closepr As Double
        Dim yearly_chg As Double
        Dim per_chg As Double
        Dim ttl_stock_value As Double
        
        'tracking of rows for opening price
        Dim Row_openpr As Double
        Row_openpr = 2
        
        'tracking of rows for summary table
        Dim Row_sum_tbl As Integer
        Row_sum_tbl = 2
        
        'Determine last row of the sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'---------------------------------------------------------------------------------------------------------------------
'LOOP THROUGH AND OUTPUT INFORMATION
'---------------------------------------------------------------------------------------------------------------------
        
        For i = 2 To lastrow
            
        'look at next row to see if it is equal to above row and we are on the same ticker.
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'store ticker symbol and print to summary table
            ticker = ws.Cells(i, 1).Value
            ws.Range("J" & Row_sum_tbl).Value = ticker
            
            'set first opening price of the ticker (starting row for each ticker)
            'select column c for opening price of this ticker
            first_openpr = ws.Cells(Row_openpr, 3).Value
                        
            'set the ending price of the ticker (current row (i), and the last row of the ticker
            last_closepr = ws.Cells(i, 6).Value
            
            'calculating chg in open and close price and print to summary table
            yearly_chg = last_closepr - first_openpr
            ws.Range("K" & Row_sum_tbl).Value = yearly_chg
            
            'calculate % Yearly Change in Stock and print to summary table
            If first_openpr = 0 Then
                per_chg = 0 'error handling if first opening price is zero
            Else
                per_chg = yearly_chg / first_openpr
            End If
           
            ws.Range("L" & Row_sum_tbl).NumberFormat = "0.00%"
            ws.Range("L" & Row_sum_tbl).Value = per_chg
            
            'calculate total stock value
            ttl_stock_value = ttl_stock_value + ws.Cells(i, 7).Value
            ws.Range("M" & Row_sum_tbl).Value = ttl_stock_value
           
 '-----------------------------------------------------------------------------------------------------------------
 'THIS IS WHERE THE HIGHLIGHTING CODE WILL GO
 'it goes here to change the color of each cell of the ticker's yearly change (variable is yearly_chg)
 'depending on whether the change is positive, negative or 0
 '-----------------------------------------------------------------------------------------------------------------
            'If yearly change is positive then color cell green
            'code for ouptut Range("K" & Row_sum_tbl).Value.Interior.ColorIndex
            If yearly_chg > 0 Then
                ws.Range("K" & Row_sum_tbl).Interior.ColorIndex = 4
            'Else if yearly change is negative then red
            ElseIf yearly_chg < 0 Then
                ws.Range("K" & Row_sum_tbl).Interior.ColorIndex = 3
            'Else if yearly change is 0 then no fill
            ElseIf yearly_chg = 0 Then
                ws.Range("K" & Row_sum_tbl).Interior.ColorIndex = 2
            End If
'-----------------------------------------------------------------------------------------------------------------
'THIS IS WHERE YOU NEED TO CHECK FOR WHAT YOU NEED TO COUNT OR RESET FOR NEXT TICKER
'The rows will need to increase for the summary table for next ticker
'the row for first opening price of for the next ticker will need to increase
'the total stock value will need to reset to 0
'-----------------------------------------------------------------------------------------------------------------
            'update the summary table row for next ticker
            Row_sum_tbl = Row_sum_tbl + 1
            
            'update the opening price row for next ticker
            Row_openpr = i + 1
            
            'reset the stock value total for next ticker
            ttl_stock_value = 0
        Else 'if the next row ticker value is the same
            'add to the stock value total
            ttl_stock_value = ttl_stock_value + ws.Cells(i, 7).Value
        End If
            
    Next i
    
'-----------------------------------------------------------------------------------------------------------------
    'THIS IS WHERE I FIND THE MIN AND MAX OF MY NEW DATA COLUMNS
    'where to start and where to finish
    'then get max or min
    'coding for function - Application.worksheetfunction.max(range("a:a"))
'-----------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'THIS IS WHERE I WILL DEFINE THE MAX OF % CHANGE INCREASE AND TOTAL STOCK VOLUME, MIN OF % CHANGE DECREASE
'---------------------------------------------------------------------------------------------------------------------

       Dim max_percent As Double
       Dim min_percent As Double
       Dim max_ttl_volume As Double
       
       Dim rw As Double
       Dim rw2 As Double
       Dim rw3 As Double
       
    'rw = Application.WorksheetFunction.Match(max_percent, Range("L:L"), 0) + Range("L:L").Row - 1
    'found the above on this website 'https://www.mrexcel.com/board/threads/how-can-i-return-the-row-number-of-a-maximum-value-in-a-range-in-vba.980012/
    
    
    'Run functions and print
    'get max percentage
    max_percent = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    rw = ws.Application.WorksheetFunction.Match(max_percent, ws.Range("L:L"), 0) + ws.Range("L:L").Row - 1
    'get min percentage
    min_percent = ws.Application.WorksheetFunction.Min(ws.Range("L:L"))
    rw1 = ws.Application.WorksheetFunction.Match(min_percent, ws.Range("L:L"), 0) + ws.Range("L:L").Row - 1
    'format and print
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("R2").Value = max_percent
    ws.Range("Q2").Value = ws.Cells(rw, 10)
    
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("R3").Value = min_percent
    ws.Range("Q3").Value = ws.Cells(rw1, 10)
     
    'get max volume total
    max_ttl_volume = ws.Application.WorksheetFunction.Max(ws.Range("M:M"))
    rw2 = ws.Application.WorksheetFunction.Match(max_ttl_volume, ws.Range("M:M"), 0) + ws.Range("M:M").Row - 1
    'print max volume total
    ws.Range("R4").Value = max_ttl_volume
    ws.Range("Q4").Value = ws.Cells(rw2, 10)
       
    'fix column with for new tables
    'ws.Worksheets("A").Columns("J:R").AutoFit
Next ws

End Sub
