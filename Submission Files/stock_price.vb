Sub StockPrice()

' Loop through each worksheet
For Each ws In Worksheets

    ' Define yearly change, % change, and total colume as variables
    Dim yr_ch As Double
    Dim pc_ch As Double
    Dim total As Double
    Dim open_price As Double
    Dim close_price As Double

    ' Set them to 0
    yr_ch = 0
    pc_ch = 0
    total = 0
    open_price = 0
    close_price = 0
    
    ' Create headers for the summary table
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly change"
    ws.Range("K1") = "Percent change"
    ws.Range("L1") = "Total volume"
    
    ' Set the summary table values to begin at row 2
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ' Define the last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Grab the open price for the first ticker
    open_price = ws.Cells(2, 3).Value
    
    ' Interate through the rows
    For j = 2 To LastRow
    
        ' If the next ticker name is different from the active cell...
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
        
        ' Output the active cell ticker name in the summary table
        ws.Cells(summary_table_row, 9).Value = ws.Cells(j, 1).Value
        
        ' Calculate and output the yearly change in the summary table
            ' Grab the closing price
            close_price = ws.Cells(j, 6).Value
            ' Subtract the opening price from the closing price to calculate yearly change
            yr_ch = close_price - open_price
        ' Output into summary table
        ws.Cells(summary_table_row, 10).Value = yr_ch
        
        ' Calculate and output the percent change in the summary table
            ' Percent change = yearly change divided by opening price
            pc_ch = (yr_ch / open_price)
        ' Output in the summary table
        ws.Cells(summary_table_row, 11).Value = pc_ch
        ' Format as a percentage
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        
        ' Add to the total volume
        total = total + ws.Cells(j, 7).Value
        
        ' Output the total in the summary table
        ws.Cells(summary_table_row, 12).Value = total
        
        ' Reset everything to 0
        yr_ch = 0
        ph_ch = 0
        total = 0
        open_price = 0
        close_price = 0

        ' Grab the new open price
        open_price = ws.Cells(j + 1, 3).Value
        
        ' Add one to the summary table row to correctly format the summary table
        summary_table_row = summary_table_row + 1
        
        ' If the next row is the same ticker name as the previous...
        Else
        ' Add to the total volume
        total = total + ws.Cells(j, 7).Value
        
        End If
        
    Next j

Next ws

End Sub
