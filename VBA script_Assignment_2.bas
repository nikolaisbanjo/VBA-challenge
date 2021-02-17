Attribute VB_Name = "Module11"
Sub stock_market()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
If ws.Name <> "instructions" Then
ws.Activate

Dim ticker_ID As String
Dim opening_price As Double
Dim closing_price As Double
Dim price_change As Double
Dim percent_change As Double
Dim total_stock_vol As Double


Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_total_volume As Double

'format Column O
Columns("O").ColumnWidth = 25

Dim LastRow As Double

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Range("I1") = "Ticker"
Range("P1") = "Ticker"
Range("J1") = "Yearly Price Change"
Range("K1") = "Percent Change (%)"
Range("L1") = "Total Stock Volume"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest total stock volume"


table_row = 2 'this is required for the code to print out all the different tickets on the rows below. Without this the code will simply print the last ticket ID and it's total volume.

total_stock_vol = 0
closing_price = 0

'Run through every row (i)
'If the cells in column 1 read the same ticker, count the total volume and print that in cell L2

opening_price = Cells(2, 3)

For i = 2 To LastRow

    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then    'if the ticker value is not the same then
        ticker_ID = Cells(i, 1).Value   'here i is the last row of the ticker_ID somehow
        
        
        closing_price = Cells(i, 6).Value
        price_change = closing_price - opening_price
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
        
        If opening_price = 0 Then
        percent_change = 0
        Else
        percent_change = Round((price_change / opening_price) * 100, [2])
        End If
        
        'redefine opening price for the next lot of data
        opening_price = Cells(i + 1, 3)
        
       'print values
        
        Range("I" & table_row) = ticker_ID
        Range("L" & table_row) = total_stock_vol
        Range("J" & table_row) = price_change
        Range("K" & table_row) = percent_change
        
        'format price change colors
           
        If price_change > 0 Then
        Range("J" & table_row).Interior.ColorIndex = 4
        
        ElseIf price_change < 0 Then
        Range("J" & table_row).Interior.ColorIndex = 3
        
        End If
    
        'use Match function to find greatest %increase, decrease and max total stock volume
        
        greatest_percent_inc = WorksheetFunction.Max(Range("K2:K" & LastRow))
        Range("Q2") = greatest_percent_inc
        ticker_max_pos = WorksheetFunction.Match(greatest_percent_inc, Range("K2:K" & LastRow), 0)
        ticker_max = Cells(ticker_max_pos, 9)
        Range("P2") = ticker_max
        
        greatest_percent_dec = WorksheetFunction.Min(Range("K2:K" & LastRow))
        Range("Q3") = greatest_percent_dec
        ticker_min_pos = WorksheetFunction.Match(greatest_percent_dec, Range("K2:K" & LastRow), 0)
        ticker_min = Cells(ticker_min_pos, 9)
        Range("P3") = ticker_min
        
        max_stock = WorksheetFunction.Max(Range("L2:L" & LastRow))
        Range("Q4") = max_stock
        ticker_maxstock_pos = WorksheetFunction.Match(max_stock, Range("L2:L" & LastRow))
        ticker_maxstock = Cells(ticker_maxstock_pos, 9)
        Range("P4") = ticker_maxstock
    
        'print values in subsequent rows
        table_row = table_row + 1
           
        'reset stock volume
        total_stock_vol = 0
         
        Else
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
                     
        End If
           
     Next i
     
End If
     
     
Next ws


End Sub

