Sub StockAnalysis()

    ' --------------------------------------------------
    ' |     CREATE AND DEFINE ALL OF THE VARIABLES     |
    ' |          THAT ARE USED IN THIS SCRIPT          |
    ' --------------------------------------------------
        
    ' Create and define variables for the row and column counters
    ' RowCount is set to 2 in the For loop that goes through all of the rows
    ' LongLong is used in case the rows exceed the limits of Long
    Dim RowCount As LongLong
    Dim ColumnCount As LongLong
    ColumnCount = 1
    
    ' Create and define variables for the ticker symbol, quarterly change, open/close prices, percent change, and total stock volume
    ' LongLong is used in case the rows exceed the limits of Long
    Dim TickerSymbol As String
    Dim QuarterlyChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As LongLong
    
    ' Create and define variable to hold the value of the last row
    ' LongLong is used in case the rows exceed the limits of Long
    Dim LastRow As LongLong
    
    ' Create and define variables for the calculations of the stocks with the greatest % increase/decrease and the greatest total volume
    Dim PercentMaxIncrease As Double
    Dim PercentMinDecrease As Double
    Dim TotalMaxVolume As LongLong
    Dim LastRowCalc As Long
    
    ' Create and define the worksheet variable as WS
    Dim WS As Worksheet
    
    ' Keep track of the location for the output of ticker/quarterly change/percent change/total stock volume table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
  
  
    ' --------------------------------------------------------------------
    ' |          LOOPS THROUGH ALL OF THE STOCKS AND OUTPUTS             |
    ' |   VARIOUS INFORMATION TO A NEW TABLE TO THE RIGHT OF THE DATA    |
    ' --------------------------------------------------------------------
    
    ' Loop through all of the sheets
    For Each WS In Worksheets
             
        ' Determine the value of the very last row in current sheet
        LastRow = WS.Cells(WS.Rows.Count, 1).End(xlUp).Row
        
        ' Get the name of the worksheet
        ' WSName = WS.Name
        
        ' Test to see if a message box will show the correct worksheet name is being retrieved
        ' MsgBox (WSName)
        
        ' Labels the appropriate column headers and rows on all sheets
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Quarterly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
                       
        ' Loop through all of the rows
        For RowCount = 2 To LastRow

            ' Search for when the value of the next cell is different than that of the current cell
            If WS.Cells(RowCount + 1, ColumnCount).Value <> WS.Cells(RowCount, ColumnCount).Value Then
                    
                ' Test to see if a message box will show the change from current cell to the next
                ' MsgBox (WS.Cells(RowCount, ColumnCount).Value & " and then " & WS.Cells(RowCount + 1, ColumnCount).Value)
                
                ' Set the value of the ticker symbol
                TickerSymbol = WS.Cells(RowCount, ColumnCount).Value
                
                ' Print the ticker symbol in the summary table
                WS.Range("I" & SummaryTableRow).Value = TickerSymbol

                ' Add to the total stock amount
                TotalStockVolume = TotalStockVolume + WS.Cells(RowCount, 7).Value
                
                ' Print the volume amount to the summary table
                WS.Range("L" & SummaryTableRow).Value = TotalStockVolume
                
                ' Set the value of the closing price
                ClosePrice = WS.Cells(RowCount, 6).Value
                                                
                ' Calculates the quarterly change by subtracting the first opening price from the last closing price for ticker
                QuarterlyChange = ClosePrice - OpenPrice
                
                ' Calculates the percent change by subtracting the first opening price from the last closing price for ticker and dividing that by the opening price
                PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                
                ' Print the values of the quartlery change and percent change to the summary table
                WS.Range("J" & SummaryTableRow).Value = QuarterlyChange
                WS.Range("K" & SummaryTableRow).Value = PercentChange

                ' Add one to the summary table row
                SummaryTableRow = SummaryTableRow + 1

                ' Reset the total stock volume
                TotalStockVolume = 0
                
                ' Formats the percentage change column to show a percentage and the quarterly change
                ' to show up to two decimal places for the "0" to show up as "0.00"
                ' WS.Range("K" & SummaryTableRow - 1).Style = "Percent"
                WS.Range("J" & SummaryTableRow - 1).NumberFormat = "0.00"
                WS.Range("K" & SummaryTableRow - 1).NumberFormat = "0.00%"
                
            ' Search for when the value of the previous cell is different than that of the current cell
            ElseIf WS.Cells(RowCount - 1, ColumnCount).Value <> WS.Cells(RowCount, ColumnCount).Value Then
            
                ' Set the value of the opening price
                OpenPrice = WS.Cells(RowCount, 3).Value
                
                ' Add to the total stock amount
                TotalStockVolume = TotalStockVolume + WS.Cells(RowCount, 7).Value
                
            ' If the cell immediately following a row is the same ticker...
            Else

                ' Add to the total stock volume
                TotalStockVolume = TotalStockVolume + WS.Cells(RowCount, 7).Value
                
            End If
                                
        Next RowCount
        
        ' Reset the summary table row
        SummaryTableRow = 2
        
        
        ' -------------------------------------------------------------------
        ' |           CALCULATIONS FOR THE TABLE CONTAINING THE             |
        ' |     GREATEST % INCREASE/DECREASE AND GREATEST TOTAL VOLUME      |
        ' -------------------------------------------------------------------
        
        ' Determine the value of the very last row in the percent change column
        LastRowCalc = WS.Cells(WS.Rows.Count, "K").End(xlUp).Row
        
        ' Find the max/largest percent value and print the value to appropriate table
        PercentMaxIncrease = WorksheetFunction.Max(WS.Range("K2:K" & LastRowCalc))
        WS.Range("Q2").Value = "%" & (PercentMaxIncrease * 100)
        
        ' Find the min/smallest percent value and print the value to appropriate table
        PercentMinDecrease = WorksheetFunction.Min(WS.Range("K2:K" & LastRowCalc))
        WS.Range("Q3").Value = "%" & (PercentMinDecrease * 100)
        
        ' Determine the value of the very last row in the total stock volume column
        LastRowCalc = WS.Cells(WS.Rows.Count, "L").End(xlUp).Row
        
        ' Find the max/largest total volume value and print the value to appropriate table
        TotalMaxVolume = WorksheetFunction.Max(WS.Range("L2:L" & LastRowCalc))
        WS.Range("Q4").Value = TotalMaxVolume
               
        ' Find the max/largest total volume value and print the ticker to appropriate cell
        MaxIncreaseTicker = WorksheetFunction.Match(WorksheetFunction.Max(WS.Range("K2:K" & LastRowCalc)), WS.Range("K2:K" & LastRowCalc), 0)
        WS.Range("P2") = WS.Cells(MaxIncreaseTicker + 1, 9)
        
        ' Find the min/smallest percent value and print the ticker to appropriate cell
        MinDecreaseTicker = WorksheetFunction.Match(WorksheetFunction.Min(WS.Range("K2:K" & LastRowCalc)), WS.Range("K2:K" & LastRowCalc), 0)
        WS.Range("P3") = WS.Cells(MinDecreaseTicker + 1, 9)
        
        ' Find the max/largest percent value and print the ticker to appropriate cell
        MaxTotalVolTicker = WorksheetFunction.Match(WorksheetFunction.Max(WS.Range("L2:L" & LastRowCalc)), WS.Range("L2:L" & LastRowCalc), 0)
        WS.Range("P4") = WS.Cells(MaxTotalVolTicker + 1, 9)
               
               
        ' -----------------------------------------------
        ' |   FORMATTING FOR AUTO-SIZING COLUMNS AND    |
        ' |     CONDITIONAL FORMATTING FOR COLUMNS      |
        ' -----------------------------------------------
        
        ' Resizes each column from I-L and O-Q to auto-fit with data in cell
        WS.Columns("I:L").AutoFit
        WS.Columns("O:Q").AutoFit
        
        ' Loop through all of the rows
        For RowCount = 2 To LastRow
        
                ' Fill in the quarterly change column with green color for positive values and red color for negative values
                If WS.Cells(RowCount, 10).Value > 0 Then
                    WS.Cells(RowCount, 10).Interior.Color = RGB(0, 255, 0)
                
                ' If the value in cell is 0, then fills in with white
                ElseIf WS.Cells(RowCount, 10).Value = 0 Then
                    WS.Cells(RowCount, 10).Interior.Color = RGB(255, 255, 255)
                
                ' Else it will fill in the rest of the values with red
                Else
                    WS.Cells(RowCount, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
        Next RowCount
           
    Next WS
    
    ' Reset the summary table row after moving on to the next sheet otherwise the ticker/quarterly change/etc.
    ' table will print on the row where the previous sheet left off
    SummaryTableRow = 2
    
End Sub

