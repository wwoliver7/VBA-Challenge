Sub Multiple_Year_StockData_WO():

    
        Dim WorksheetName As String
        '
        Dim i As Long
        Dim j As Long
        Dim TickerCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentChange As Double
        Dim GIncrease As Double
        Dim GDecrease As Double
        Dim GTotalVolume As Double
        
        TickerCount = 2
        j = 2
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        LastRowA = Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
            
                
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                
                Cells(TickerCount, 9).Value = Cells(i, 1).Value
                
                
                Cells(TickerCount, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
                
                    
                    If Cells(TickerCount, 10).Value < 0 Then
                
                  
                    Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    
                    Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    
                    If Cells(j, 3).Value <> 0 Then
                    PercentChange = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
                    
                    
                    Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                    
                    
                    Else
                    
                    Cells(TickerCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                
                 Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(Cells(j, 7), Cells(i, 7)))
                
                
                TickerCount = TickerCount + 1
                j = i + 1
                
                End If
            
            Next i
            
        
        LastRowI = Cells(Rows.Count, 9).End(xlUp).Row
        
     
        GIncrease = Cells(2, 11).Value
        GDecrease = Cells(2, 11).Value
        GTotalVolume = Cells(2, 12).Value
            
            For i = 2 To LastRowI
            
                
                If Cells(i, 12).Value > GTotalVolume Then
                GTotalVolume = Cells(i, 12).Value
                Cells(4, 16).Value = Cells(i, 9).Value
                
                Else
                
                GTotalVolume = GTotalVolume
                
                End If
                
                
                If Cells(i, 11).Value > GIncrease Then
                GIncrease = Cells(i, 11).Value
                Cells(2, 16).Value = Cells(i, 9).Value
                
                Else
                
                GIncrease= GIncrease
                
                End If
                
                
                If Cells(i, 11).Value < GDecrease Then
                GDecrease = Cells(i, 11).Value
                Cells(3, 16).Value = Cells(i, 9).Value
                
                Else
                
                GDecrease = GDecrease
                
                End If
                
            
            Cells(2, 17).Value = Format(GIncrease, "Percent")
            Cells(3, 17).Value = Format(GDecrease, "Percent")
            Cells(4, 17).Value = Format(GTotalVolume, "Scientific")
            
            Next i
            
        
        
            
End Sub
