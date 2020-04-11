Sub stockmarket():

    'Loop through all the sheets
    For Each ws In Worksheets
    
        'Set variable for holding ticker
        Dim Ticker_Sym As String
        
        'Set variable to grab open and close
        Dim Open_Stock As Double
        Dim Close_Stock As Double
        
        'Set variable for holding volume total
        Dim Volume_Total As Double
        Volume_Total = 0
        
        'Keep track of location for summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Set variable to track ticker set
        Dim Ticker_Array As Double
        Row_Count = 0
        
        'Add headers for summary table
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'Loop through all stock data
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
             
            'If cell immediately following a row is the same ticker
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            
                'Start row count
                Row_Count = Row_Count + 1

                'Add to the volume total
                Volume_Total = (Volume_Total) + (ws.Cells(i, 7).Value)
        
            'Check if we're still within same ticker, if it is not
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                'Grab ticker symbol and print on summary table
                Ticker_Sym = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Sym
                
                'Grab open value
                Open_Stock = ws.Cells(i - (Row_Count), 3).Value
                
                'Grab close value
                Close_Stock = ws.Cells(i, 6).Value
                    
                'Calculate yearly change and print on summary table
                ws.Range("J" & Summary_Table_Row).Value = (Close_Stock - Open_Stock)
                
                'Format color for yearly change column
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                    ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                End If
                    
                'Calculate yearly change percentage and print on summary table
                If Open_Stock <> 0 Then
                    ws.Range("K" & Summary_Table_Row).Value = (Close_Stock - Open_Stock) / Open_Stock
                    'Else when Open_Stock is 0 then find next non-zero value
                    ElseIf Open_Stock = 0 Then
                        For j = i - (Row_Count) To i
                            If ws.Cells(j, 3).Value <> 0 Then
                                Open_Stock = ws.Cells(j, 3).Value
                                ws.Range("K" & Summary_Table_Row).Value = (Close_Stock - Open_Stock) / Open_Stock
                                Exit For
                            End If
                        Next j
                        
                End If
                    
                'Style yearly change as percentage
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Set volume total and print on summary table
                Volume_Total = (Volume_Total) + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Volume_Total
                        
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                    
                'Reset volume total and row count
                Volume_Total = 0
                Row_Count = 0
                
            End If
            
        Next i
    
        'Summary table for greatest % increase, greatest % decrease and greatest total volume
        
            'Add column and row headers for summary table
            ws.Range("O1") = "Ticker"
            ws.Range("P1") = "Value"
            ws.Range("N2") = "Greatest % Increase"
            ws.Range("N3") = "Greatest % Decrease"
            ws.Range("N4") = "Greatest Total Volume"
            
            'Set for loop
            lastRowX = Cells(Rows.Count, 9).End(xlUp).Row
            For k = 2 To lastRowX
            
            'Set conditionals to grab and print greatest % increase, % decrease and total volume
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRowX)) Then
                ws.Range("O2").Value = ws.Cells(k, 9).Value
                ws.Range("P2").Value = ws.Cells(k, 11).Value
                ws.Range("P2").NumberFormat = "0.00%"
                
                ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRowX)) Then
                ws.Range("O3").Value = ws.Cells(k, 9).Value
                ws.Range("P3").Value = ws.Cells(k, 11).Value
                ws.Range("P3").NumberFormat = "0.00%"
                
                ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRowX)) Then
                ws.Range("O4").Value = ws.Cells(k, 9).Value
                ws.Range("P4").Value = ws.Cells(k, 12).Value
                
            End If
    
            Next k
                      
        'Autofit columns on worksheet
        ws.Columns.AutoFit
        
    Next ws
    
End Sub