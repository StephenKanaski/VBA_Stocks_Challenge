Sub Stock_Analysis_test()
    ' loop through all worksheets
    ' output Ticker Symbol, Yearly Change, % Yearly Change, Total Stock Volume
    ' Conditional Formatting, green for positive change and red for negative change
    
    ' add new worksheet and move to first position in workbook
    'Sheets.Add.Name = "Combined_Stock_Data"
    
    ' move Combined_Stock_Data sheet to beginning of workbook
    'Sheets("Combined_Stock_Data").Move Before:=Sheets(1)
    
    ' specify location of the combined data sheet
    'Set Combined_Sheet = Worksheets("Combined_Stock_Data")
    
    ' Set Names to Column Headers in Combined_Sheet
    'Combined_Sheet.Cells(1, 1).Value = "Ticker"
    'Combined_Sheet.Cells(1, 2).Value = "Yearly_Change"
    'Combined_Sheet.Cells(1, 3).Value = "Percent_Change"
    'Combined_Sheet.Cells(1, 4).Value = "Total_Stock_Volume"
    ' format to fit data
    'Combined_Sheet.Columns.AutoFit
    
For Each ws In Worksheets
    ws.Activate
    
    
    ' set variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    Dim Total_Stock_Volume As Long
    Total_Stock_Volume = 0
    
    Dim open_price As Double
    open_price = ws.Cells(2, 3).Value

    Dim close_price As Double
        
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Label Summary Table Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Trade Volume"
    
    ' Determine LastRow
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Determine LastColumn
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' retrieve and store opening, closing, and volume values
    
    
    For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            
            ' assign closing price value to run yearly change and percent change
            close_price = ws.Cells(i, 6).Value
            Yearly_Change = (close_price - open_price)
            Percent_Change = Yearly_Change / open_price
            
            ' Determine trade volume for each ticker
            Total_Stock_Volume = stock_volume + ws.Cells(i, 7).Value
            
            ' Print retrieved data to summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            ' omit non-divisible, undefined calculations
            If open_price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / open_price
            End If
            
            
            
        ' add one to Combined_Sheet data
            Summary_Table_Row = Summary_Table_Row + 1
            
        ' reset total stock volume
            Total_Stock_Volume = 0
        
        Else
            Total_Stock_Volume = stock_volume + ws.Cells(i, 7).Value
            
              
        
        End If
    
    Next i

    
    LastRow_Summary_Table = ws.Cells(Rows.Count, 4).End(xlUp).Row
    
    ' Set color formatting (green for positive change and red for negative change)
    For i = 2 To LastRow_Summary_Table
        
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    Next i
    
    ' loop through summary table to find Greatest Percent Increase,
    ' Greatest Percent Decrease, and Greatest Total Volume
    
    'Define Row Titles and Column Headers for max,min table
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value Figure"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    For i = 2 To LastRow_Summary_Table
    
        ' Greatest Percent Increase
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & LastRow_Summary_Table)) Then
            ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
            ' format value figure cell as percentage
            ws.Cells(2, 16).NumberFormat = "0.00%"
            
        ' Greatest Percent Decrease
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & LastRow_Summary_Table)) Then
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
            ' format value figure cell as percentage
            ws.Cells(3, 16).NumberFormat = "0.00%"
            
        ' Greatest Total Volume
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & LastRow_Summary_Table)) Then
            ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
            
        End If
    
    Next i
    
    
    Next ws
    
End Sub

