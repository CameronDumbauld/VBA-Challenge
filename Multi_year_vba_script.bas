Attribute VB_Name = "Module1"
Sub yearlydata():

    For Each ws In Worksheets
    ' setting code to work across all worksheets

        Dim ticker As String
        Dim ticker_total As Double
        Dim percent_change_total As Double
        Dim yearly_change_total As Double
        
        ticker_total = 0
        percent_change_total = 0
        yearly_change_total = 0
        
        summary_table_row = 2
        
        ' Setting string values to reflect data in each column
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Grestest Total Volume"
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "Value"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' setting range to find the first and last row of each worksheet. I didn't each ws had a different number of rows at frst and had to adjust my code accordingly.
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ' setting code to identify each different ticker label
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ' This sums the total volume for each stock ticker
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            ws.Range("K" & summary_table_row).Value = ticker
            ws.Range("L" & summary_table_row).Value = ticker_total
            summary_table_row = summary_table_row + 1
            ticker_total = 0
            
            Else
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
            End If
            
            ' The next three if statements calculate the yearly change in each ticker for each year.
            ' Getting these to work for each year was tricky for me. I ended up having to adjust the 'left' value for 2018 and 2019 because the code was initially only working on the 2020 worksheet
            ' I took advantage of the fact that within each worksheet, there an identicle number of rows for each ticker allowing the subtraction-from-i value to work throughout the entire sheet.
            ' Originally, the values for the yearly change column were appearing in every other cell. Subtracting 1 from the summary table value corrected this.
            ' The 'right' function value (12/31) makes sure the code is subtracting the correct sumber of cells from only the last row for each ticker
            
            
            If Right(ws.Cells(i, 2).Value, 4) = 1231 And Left(ws.Cells(i, 2).Value, 4) = 2018 Then
            percent_change_total = percent_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 250, 3).Value) / ws.Cells(i - 250, 3).Value
            yearly_change_total = yearly_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 250, 3).Value)
            ws.Range("m" & summary_table_row - 1).Value = percent_change_total
            ws.Range("n" & summary_table_row - 1).Value = yearly_change_total
            yearly_change_total = 0
            percent_change_total = 0
            
            End If
            
            
            If Right(ws.Cells(i, 2).Value, 4) = 1231 And Left(ws.Cells(i, 2).Value, 4) = 2019 Then
            percent_change_total = percent_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 251, 3).Value) / ws.Cells(i - 251, 3).Value
            yearly_change_total = yearly_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 251, 3).Value)
            ws.Range("m" & summary_table_row - 1).Value = percent_change_total
            ws.Range("n" & summary_table_row - 1).Value = yearly_change_total
            yearly_change_total = 0
            percent_change_total = 0
            
            End If
            
                
            If Right(ws.Cells(i, 2).Value, 5) = 1231 And Left(ws.Cells(i, 2).Value, 4) = 2020 Then
            percent_change_total = percent_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 252, 3).Value) / ws.Cells(i - 252, 3).Value
            yearly_change_total = yearly_change_total + (ws.Cells(i, 6).Value - ws.Cells(i - 252, 3).Value)
            ws.Range("m" & summary_table_row - 1).Value = percent_change_total
            ws.Range("n" & summary_table_row - 1).Value = yearly_change_total
            yearly_change_total = 0
            percent_change_total = 0
            
            
            End If
            
            
            ' These next lines identify the minimum and maximum values for percent change and the maximum value for total volume
            ' The following if statements dictate that if a value is indeed one of the needed maximums or minimums, it should be displayed in the appropriate cells
            ' This was a section I anticpated to be tricky, but the owrksheet function made everything very smooth (thank goodness)
            ws.Cells(2, 19).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & ws.Cells(Rows.Count, 1).End(xlUp).Row))
            ws.Cells(3, 19).Value = Application.WorksheetFunction.Min(ws.Range("m2:m" & ws.Cells(Rows.Count, 1).End(xlUp).Row))
            ws.Cells(4, 19).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ws.Cells(Rows.Count, 1).End(xlUp).Row))
            
            
            If ws.Cells(i, 13).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & ws.Cells(Rows.Count, 1).End(xlUp).Row)) Then
            ws.Cells(2, 18).Value = ws.Cells(i, 11).Value
            
            End If
            
            If ws.Cells(i, 13).Value = Application.WorksheetFunction.Min(ws.Range("m2:m" & ws.Cells(Rows.Count, 1).End(xlUp).Row)) Then
            ws.Cells(3, 18).Value = ws.Cells(i, 11).Value
            
            End If
            
            If ws.Cells(i, 12) = Application.WorksheetFunction.Min(ws.Range("l2:l" & ws.Cells(Rows.Count, 1).End(xlUp).Row)) Then
            ws.Cells(4, 18).Value = ws.Cells(i, 11).Value
            
            End If
            
            
        Next i
        
    Next ws

End Sub


