Attribute VB_Name = "Module1"
Sub Mulitple_Year_Stock_Data()

    ' Identify lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set variable for ticker ID
    Dim ticker_id As String
    
    ' Set a variable for the yearly change
    Dim yearly_change As Double
    
    ' Set variable for opening/closing price
    Dim opening_price As Double
    Dim closing_price As Double
    
    ' Set variable for percent change
    Dim percent_change As Double
    
    ' Set initial variable for total stock
    Dim stock_total As Variant
    stock_total = 0
    
    ' Identify first/last row for each ticker ID
    Dim firstrow As Boolean
    firstrow = True
    
    ' Keep track of location of ticker ID in a summary table
    Dim Summary_table_row As Integer
    Summary_table_row = 2
    
    ' Loop through yearly change
    For i = 2 To lastrow
    
        If firstrow = True Then
        opening_price = Cells(i, 3).Value
        firstrow = False
    
        End If
    
        stock_total = stock_total + Cells(i, 7).Value
    
        ' Identify when the ticker ID changes
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            ' Set the ticker ID
            ticker_id = Cells(i, 1).Value
            
            closing_price = Cells(i, 6).Value
            
            firstrow = True
            
            ' Calculate yearly change
            yearly_change = closing_price - opening_price
            
            ' Calculate percet change
            percent_change = (closing_price - opening_price) / opening_price * 100
            
            ' Print the ticker ID in the summary table
            Range("i" & Summary_table_row).Value = ticker_id
            
            ' Print yearly change in summary table
            Range("j" & Summary_table_row).Value = yearly_change
            
            ' Print percent change in summary table
            Range("k" & Summary_table_row).Value = percent_change
            
            ' Print stock total in summary table
            Range("l" & Summary_table_row).Value = stock_total
            
            ' Add conditional formatting in J column
            If Cells(Summary_table_row, 10).Value >= 0 Then
            
            ' Assign color index
            Cells(Summary_table_row, 10).Interior.ColorIndex = 4
            
            ' If less than 0
            ElseIf Cells(Summary_table_row, 10).Value < 0 Then
            
            ' Assign color index
            Cells(Summary_table_row, 10).Interior.ColorIndex = 3
            
            End If
            
            ' Add conditional formatting in K column
            If Cells(Summary_table_row, 11).Value >= 0 Then
            
            ' Assign color index
            Cells(Summary_table_row, 11).Interior.ColorIndex = 4
            
            ' If less than 0
            ElseIf Cells(Summary_table_row).Value < 0 Then
            
            ' Assign color index
            Cells(Summary_table_row, 11).Interior.ColorIndex = 3
            
            End If
            
            ' Add one to the summary table row
            Summary_table_row = Summary_table_row + 1
            
            ' Reset stock total counter
            stock_total = 0
            
        End If
        
    Next i
    
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
        inc_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
        Range("P4") = Cells(inc_index + 1, 9)
        
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
        inc_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        Range("P2") = Cells(inc_index + 1, 9)
        
        Range("Q3") = "%" & Worksheet.Function.Min(Range("K2:K" & lastrow)) * 100
        inc_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        Range("P3") = Cells(inc_index + 1, 9)
        
        
End Sub

