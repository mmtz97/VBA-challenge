Sub StockData():

    For Each ws In Worksheets
    
        Dim WorksheetN As String
    
        Dim PercentChange As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim GreatVolume As Double
        Dim a As Long
        Dim b As Long
        Dim TickCount As Long
        Dim RowA As Long
        Dim RowI As Long

        
        WorksheetN = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        TickCount = 2
        
        b = 2
        
        RowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For a = 2 To RowA
            

                If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
                
                ws.Cells(TickCount, 9).Value = ws.Cells(a, 1).Value
                
                ws.Cells(TickCount, 10).Value = ws.Cells(a, 6).Value - ws.Cells(b, 3).Value
                
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(b, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(a, 6).Value - ws.Cells(b, 3).Value) / ws.Cells(b, 3).Value)
                    
                    ws.Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(b, 7), ws.Cells(a, 7)))
                
                TickCount = TickCount + 1
                
                b = a + 1
                
                End If
            
            Next a
            
        RowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            For a = 2 To RowI
            
                If ws.Cells(a, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(a, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                

                If ws.Cells(a, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(a, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                If ws.Cells(a, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(a, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(a, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next a
            
        Worksheets(WorksheetN).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
