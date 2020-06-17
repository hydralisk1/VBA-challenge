Sub stock_data()

    ' Variable declaration
    Dim last_row As Double          ' the last row of worksheets
    Dim pointer As Double           ' summary table row pointer
    Dim open_price_row As Double    ' opening price row pointer
    Dim total_vol As Double         ' total stock volume
    
    total_vol = 0

    ' For statement to process all the worksheets
    For Each ws In Worksheets
    
        ' lables for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' getting the ends of the row numbers on each sheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' summary table's staring row pointer
        pointer = 2
        
        ' first opening price row pointer
        open_price_row = 2
                
        ' loop to the end of the rows
        For i = 2 To last_row
            ' accumulating total stock volumes
            total_vol = total_vol + ws.Cells(i, 7).Value
            
            ' if the next ticker is different from the current ticker, put values in the summary table
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(pointer, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(pointer, 10).Value = ws.Cells(i, 6).Value - ws.Cells(open_price_row, 3).Value
                
                ' in order to check if opening price is zero because nothing can be divided by zero
                If ws.Cells(open_price_row, 3).Value = 0 Then
                    ws.Cells(pointer, 11).Value = 0
                Else
                    ws.Cells(pointer, 11).Value = ws.Cells(pointer, 10).Value / ws.Cells(open_price_row, 3).Value
                End If
                
                ws.Cells(pointer, 12).Value = total_vol
                
                ' if increased, filling with green, if decreased,filling with red
                If ws.Cells(pointer, 10).Value < 0 Then
                    ws.Cells(pointer, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(pointer, 10).Value > 0 Then
                    ws.Cells(pointer, 10).Interior.ColorIndex = 4
                End If
                
                ' making cell format %
                ws.Cells(pointer, 11).NumberFormat = "0.00%"
                
                
                ' initializing total stock volume
                total_vol = 0
                
                ' increasing summary table's row pointer
                pointer = pointer + 1
                
                ' next opening price pointer
                open_price_row = i + 1
            End If
        Next i
        
        ' adjusting column sizes automatically
        ws.Columns("A:L").AutoFit
    
    Next ws
End Sub
