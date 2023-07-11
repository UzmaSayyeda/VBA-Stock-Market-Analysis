Sub StockAnalysis():
  
    'loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
    
        'declaring variables
        Dim OldTicker As String
        
        Dim NewTicker As String
        
        Dim TotalStockVolume As Double
        
        Dim OpenYear As Double
        
        Dim CloseYear As Double
        
        Dim YearlyChange As Double
        
        Dim PercentChange As Double
        
        Dim y As Double
        
        Dim i As Double
        
            'insert headers
            
            ws.Range("I1").Value = "Ticker"
            
            ws.Range("J1").Value = "Yearly_change"
            
            ws.Range("K1").Value = "Percent_change"
            
            ws.Range("L1").Value = "Total_Stock_volume"
            
            ws.Range("O2").Value = "Greatest % Increase"
            
            ws.Range("O3").Value = "Greatest % Decrease"
            
            ws.Range("O4").Value = "Greatest Total Volume"
            
            ws.Range("P1").Value = "Ticker"
            
            ws.Range("Q1").Value = "Value"
            
            ws.Range("I1:q4").Columns.AutoFit
    
                    summary_table_row = 2
                    
                    'to find the last row
                    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                    
                    y = 2
                    
                         ' loop through each row
                         For i = 2 To LastRow
                         
                            
                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            
                            NewTicker = ws.Cells(i, 1).Value
                            
                            'define range for OpenYear
                            OpenYear = ws.Cells(y, 3).Value
                            
                            'define range CloseYear
                            CloseYear = ws.Cells(i, 6).Value
                            
                            'calculate yearly change
                            YearlyChange = CloseYear - OpenYear
                            
                            'calculate PercentChnage
                            PercentChange = YearlyChange / ws.Cells(y, 3)
                            
                            'Total Volume
                            TotalStockVolume = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(i, 7)))
                            
                            'insert values
                            ws.Range("I" & summary_table_row).Value = NewTicker
                            
                            ws.Range("L" & summary_table_row).Value = TotalStockVolume
                            
                            ws.Range("j" & summary_table_row).Value = YearlyChange
                            
                            ws.Range("k" & summary_table_row).Value = PercentChange
                            
                            summary_table_row = summary_table_row + 1
                            
                            y = (i + 1)
                     
                            End If
                                
                        Next i
                                
                            'format %
                            ws.Columns("K").NumberFormat = "0.00%"
            
                                'colour conditioning
                                For j = 2 To LastRow
                                
                                    If ws.Cells(j, 10).Value > 0 Then
                                        ws.Cells(j, 10).Interior.Color = vbGreen
                                    Else
                                        ws.Cells(j, 10).Interior.Color = vbRed
  
                                    End If
                                    
                                        If ws.Cells(j, 11).Value > 0 Then
                                            ws.Cells(j, 11).Interior.Color = vbGreen
                                        Else
                                            ws.Cells(j, 11).Interior.Color = vbRed
                                            
                                        End If
                                 Next j
                                
              'Functionality
              'variables
              Dim GreatestIncrease
              
              Dim GreatestDecrease
              
              Dim GreatestVolume
              
              GreatestIncrease = ws.Range("K:K").Value
              
              GreatestDecrease = ws.Range("K:K").Value
              
              GreatestVolume = ws.Range("L:L").Value
              
              'find max and min
              ws.Range("Q2").Value = WorksheetFunction.Max(GreatestIncrease)
              
              ws.Range("Q3").Value = WorksheetFunction.Min(GreatestDecrease)
              
              ws.Range("Q4").Value = WorksheetFunction.Max(GreatestVolume)
              
              ws.Range("Q2:Q3").NumberFormat = "0.00%"
              
              'match values to ticker name
              MatchRowIncrease = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:k"), 0)
              
              ws.Range("P2").Value = ws.Cells(MatchRowIncrease, "I").Value
              
              MatchRowDecrease = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0)
              
              ws.Range("P3").Value = ws.Cells(MatchRowDecrease, "I").Value
              
              MatchRowVolume = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0)
              
              ws.Range("P4").Value = ws.Cells(MatchRowVolume, "I").Value
              
              
              
        Next ws
    
End Sub