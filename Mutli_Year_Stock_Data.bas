Attribute VB_Name = "Module1"
Sub MYSD():

    'establishing all variables
    Dim wkshtName As String
    Dim i As Long
    Dim j As Long
    Dim tickerCt As Long
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim percentChange As Double
    Dim grtPerIncr As Double
    Dim grtPerDecr As Double
    Dim grtTotalVol As Double


    For Each wksht In Worksheets
        
        wkshtName = wksht.Name
        
        'first row, starting row
        tickerCt = 2
        j = 2
        
        'last cell w value in col A
        lastRow = wksht.Cells(Rows.Count, 1).End(xlUp).Row
        
            'for loop going through all rows
            For i = 2 To lastRow
                'check ticker name
                If wksht.Cells(i + 1, 1).Value <> wksht.Cells(i, 1).Value Then
                
                wksht.Cells(tickerCt, 9).Value = wksht.Cells(i, 1).Value
                
                'Calculate the yearly change
                wksht.Cells(tickerCt, 10).Value = wksht.Cells(i, 6).Value - wksht.Cells(j, 3).Value
                
                    If wksht.Cells(tickerCt, 10).Value < 0 Then
                
                    'changing cell color to red
                    wksht.Cells(tickerCt, 10).Interior.ColorIndex = 3
                
                        Else
                
                    'changing cell color to green
                    wksht.Cells(tickerCt, 10).Interior.ColorIndex = 4
                
                            End If
                    
                    'Calculate the percent change
                    If wksht.Cells(j, 3).Value <> 0 Then
                    percentChange = ((wksht.Cells(i, 6).Value - wksht.Cells(j, 3).Value) / wksht.Cells(j, 3).Value)
                    
                    wksht.Cells(tickerCt, 11).Value = Format(percentChange, "Percent")
                    
                        Else
                    
                    wksht.Cells(tickerCt, 11).Value = Format(0, "Percent")
                    
                            End If
                    
                'Calculate the total volume
                wksht.Cells(tickerCt, 12).Value = WorksheetFunction.sum(Range(wksht.Cells(j, 7), wksht.Cells(i, 7)))
                
                tickerCt = tickerCt + 1
                
                'new starting row
                j = i + 1
                
                End If
            
            Next i
            
    'last row, last column
        lastRow2 = wksht.Cells(Rows.Count, 9).End(xlUp).Row
        
    grtPerIncr = wksht.Cells(2, 11).Value
    grtPerDecr = wksht.Cells(2, 11).Value
    grtTotalVol = wksht.Cells(2, 12).Value
        
            'loop for calculated values
            For i = 2 To lastRow2
            
                If wksht.Cells(i, 12).Value > grtTotalVol Then
                grtTotalVol = wksht.Cells(i, 12).Value
                wksht.Cells(4, 16).Value = wksht.Cells(i, 9).Value
                
                    Else
                
                grtTotalVol = grtTotalVol
                
                        End If
                
                If wksht.Cells(i, 11).Value > grtPerIncr Then
                grtPerIncr = wksht.Cells(i, 11).Value
                wksht.Cells(2, 16).Value = wksht.Cells(i, 9).Value
                
                    Else
                GreatIncr = GreatIncr
                
                        End If
                
                If wksht.Cells(i, 11).Value < grtPerDecr Then
                grtPerDecr = wksht.Cells(i, 11).Value
                wksht.Cells(3, 16).Value = wksht.Cells(i, 9).Value
                
                    Else
                GreatDecr = GreatDecr
                
                        End If
                
        wksht.Cells(2, 17).Value = Format(grtPerIncr, "Percent")
        wksht.Cells(3, 17).Value = Format(grtPerDecr, "Percent")
        wksht.Cells(4, 17).Value = Format(grtTotalVol, "Scientific")
            
            Next i
            
        Worksheets(wkshtName).Columns("A:Z").AutoFit
            
    Next wksht
        
End Sub
