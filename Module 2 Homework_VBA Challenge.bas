Attribute VB_Name = "Module1"
Sub Stocks()

    
    For Each ws In Worksheets
    
    Dim priceDiff As Double
    Dim total_stock As Double
    
    output_row = 2
    total_stock = 0
    
    priceDiff = 0
    opensum = ws.Cells(2, "C").Value
    
        For i = 2 To 753001
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                total_stock = total_stock + ws.Cells(i, "G").Value
                
                ' Name of stock
                ws.Cells(output_row, "J").Value = ws.Cells(i, 1).Value
                ' Total Stock Volume
                ws.Cells(output_row, "M").Value = total_stock
                ' Update row value for the next stock
                
                
                closesum = closesum + ws.Cells(i, 6).Value
                priceDiff = closesum - opensum
                
                ws.Cells(output_row, "L").Value = priceDiff / opensum * 100 & "%"
                ws.Cells(output_row, "K").Value = priceDiff
                
                total_stock = 0
                opensum = ws.Cells(i + 1, "C").Value
                closesum = 0
                
                
                output_row = output_row + 1
                
            Else
                total_stock = total_stock + ws.Cells(i, "G").Value
                    
            End If
                    
            'Conditional Formatting (Red/Green cell color) for Yearly Change
            If ws.Cells(i, "K").Value < 0 Then
                ws.Cells(i, "K").Interior.ColorIndex = 3
                
            ElseIf ws.Cells(i, "K").Value > 0 Then
                ws.Cells(i, "K").Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, "K").Interior.ColorIndex = xlNone
            End If
            
        
            
        Next i
        
    Next ws
End Sub

Sub PercentIncrease()
    For Each ws In Worksheets
    
    MaxPercent = ws.Cells(2, "L").Value
    stockname = ws.Cells(2, "J").Value
    
        For i = 2 To 3001
            If ws.Cells(i, "L").Value > MaxPercent Then
                MaxPercent = ws.Cells(i, "L").Value
                stockname = ws.Cells(i, "J").Value
            End If
            
        Next i
    ws.Cells(2, "Q").Value = stockname
    ws.Cells(2, "R").Value = MaxPercent
    Next ws
End Sub
Sub PercentDecrease()
    For Each ws In Worksheets
    
    MinPercent = ws.Cells(2, "L").Value
    stockname = ws.Cells(2, "J").Value
        For i = 2 To 3001
            If ws.Cells(i, "L").Value < MinPercent And ws.Cells(i, "L").Value < 0 Then
                MinPercent = ws.Cells(i, "L").Value
                stockname = ws.Cells(i, "J").Value
            End If
        Next i
    ws.Cells(3, "Q").Value = stockname
    ws.Cells(3, "R").Value = MinPercent * 100 & "%"
    Next ws
End Sub
Sub GreatestStock()

    For Each ws In Worksheets
    
    maxVolume = ws.Cells(2, "M").Value
    stockname = ws.Cells(2, "J").Value
    
        For i = 2 To 3001
            If ws.Cells(i, "M").Value > maxVolume Then
                maxVolume = ws.Cells(i, "M").Value
                stockname = ws.Cells(i, "J").Value
            End If
        Next i
    ws.Cells(4, "Q").Value = stockname
    ws.Cells(4, "R").Value = maxVolume
    Next ws
End Sub
