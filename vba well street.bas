Attribute VB_Name = "Module1"
Sub Wall_Street()
    
    Dim lastRow, stockVolume, greatestTotalVolume, greatestPercentageIncrease, greatestPercentageDecrease, trickerCount, count As Long
    Dim openingPrice, ClosingPrice, changePrice, percentageChange As Double
    Dim greatestVolTicker, increaseTicker, decreaseTicker As String
    
    
    greatestTotalVolume = 0
    greatestPercentageIncrease = 0
    greatestPercentageDecrease = 0
    
    For Each ws In Worksheets
        lastRow = ws.Range("A" & Rows.count).End(xlUp).Row
    
        ' initialization section
        
        stockVolume = 0
        ws.Range("J1").Value = "ticker"
        ws.Range("K1").Value = "Total Stock Volumn"
        ws.Range("L1").Value = "Average Yearly Change in Price"
        ws.Range("M1").Value = "% Yearly Change in Price"
        count = 1
        trickerCount = 0
        changePrice = 0
        TotalClosingPrice = 0
        averageChangePrice = 0
        percentage = 0
        
        For i = 2 To lastRow + 1
        
            ' one year of stock data for each run and return the total volume each stock had over that year
            
            If ws.Range("A" & i).Value = ws.Range("A" & i + 1).Value Then
                stockVolume = stockVolume + ws.Range("g" & i)
                trickerCount = trickerCount + 1
                ClosingPrice = ws.Range("f" & i).Value
                openingPrice = ws.Range("c" & i).Value
                changePrice = changePrice + ClosingPrice - openingPrice
                TotalClosingPrice = TotalClosingPrice + ClosingPrice
            Else
                stockVolume = stockVolume + ws.Range("g" & i)
                trickerCount = trickerCount + 1
                ClosingPrice = ws.Range("f" & i).Value
                openingPrice = ws.Range("c" & i).Value
                changePrice = changePrice + ClosingPrice - openingPrice
                TotalClosingPrice = TotalClosingPrice + ClosingPrice
                If trickerCount <> 0 Then
                    averageChangePrice = changePrice / trickerCount
                End If
                If TotalClosingPrice <> 0 Then
                    percentage = (changePrice * 100) / TotalClosingPrice
                End If
                count = count + 1
                ws.Range("j" & count) = ws.Range("A" & i).Value
                ws.Range("k" & count) = stockVolume
                ws.Range("l" & count) = averageChangePrice
                ws.Range("m" & count) = percentage
                
                ' conditional formatting via color
                
                If averageChangePrice < 0 Then
                    ws.Range("l" & count).Interior.ColorIndex = 3
                Else
                    ws.Range("l" & count).Interior.ColorIndex = 4
                End If
                
        
                ' greatest total volume

                If greatestTotalVolume > stockVolume Then
                    greatestTotalVolume = greatestTotalVolume
                Else
                    greatestTotalVolume = 0
                    greatestTotalVolume = stockVolume
                    greatestVolTicker = ws.Range("A" & i).Value
                End If
                
                ' greatest percentage increase
                
                If greatestPercentageIncrease > percentage Then
                    greatestPercentageIncrease = greatestPercentageIncrease
                Else
                    greatestPercentageIncrease = percentage
                    increaseTicker = ws.Range("A" & i).Value
                End If
                
                ' greatest percentage decrease
                
                If greatestPercentageDecrease < percentage Then
                    greatestPercentageDecrease = greatestPercentageDecrease
                Else
                    greatestPercentageDecrease = percentage
                    decreaseTicker = ws.Range("A" & i).Value
                End If
                
                
                
                ' after setting for first ticker , initialize the values to default
                stockVolume = 0
                trickerCount = 0
                ClosingPrice = 0
                openingPrice = 0
                changePrice = 0
                TotalClosingPrice = 0
                percentage = 0
            End If
            
        Next i
    Next ws
    
    ' create the final table
    
    Sheet1.Range("N" & 2).Value = "Greatest % increase"
    Sheet1.Range("N" & 3).Value = "Greatest % decrease"
    Sheet1.Range("N" & 4).Value = "Greatest Total Volume"
    Sheet1.Range("O" & 1).Value = "Ticker"
    Sheet1.Range("O" & 2).Value = greatestVolTicker
    Sheet1.Range("O" & 3).Value = increaseTicker
    Sheet1.Range("O" & 4).Value = decreaseTicker
    Sheet1.Range("P" & 1).Value = "Value"
    Sheet1.Range("P" & 2).Value = greatestPercentageIncrease
    Sheet1.Range("P" & 3).Value = greatestPercentageDecrease
    Sheet1.Range("P" & 4).Value = greatestTotalVolume
    
End Sub

