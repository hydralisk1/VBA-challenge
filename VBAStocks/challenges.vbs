Sub challenges()

    ' Variable declaration
    Dim last_row As Double          ' the last row for summary table
    Dim great_inc As Double         ' greatest % increase
    Dim great_dec As Double         ' greatest % decrease
    Dim great_tot_vol As Double     ' greatest total volume
    Dim great_inc_tic As String     ' ticker of greatest % increase
    Dim great_dec_tic As String     ' ticker of greatest % decrease
    Dim great_tot_vol_tic As String ' ticker of greatest total volume
      
    ' For statement to process all the worksheets
    For Each ws In Worksheets
        
        ' initializing variables
        great_inc = 0
        great_dec = 0
        great_tot_vol = 0
    
        ' getting the ends of the row numbers for summary table on each sheet
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To last_row
        
            ' getting greatest % increase
            If ws.Cells(i, 11).Value > great_inc Then
                great_inc = ws.Cells(i, 11).Value
                great_inc_tic = ws.Cells(i, 9).Value
            End If
            
            ' getting greatest % decrease
            If ws.Cells(i, 11).Value < great_dec Then
                great_dec = ws.Cells(i, 11).Value
                great_dec_tic = ws.Cells(i, 9).Value
            End If
            
            ' getting greatest total stock volume
            If ws.Cells(i, 12).Value > great_tot_vol Then
                great_tot_vol = ws.Cells(i, 12).Value
                great_tot_vol_tic = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' making table for the challenges
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greated Total Volume"
        ws.Range("P2").Value = great_inc_tic
        ws.Range("Q2").Value = great_inc
        ws.Range("P3").Value = great_dec_tic
        ws.Range("Q3").Value = great_dec
        ws.Range("P4").Value = great_tot_vol_tic
        ws.Range("Q4").Value = great_tot_vol
        
        ' making cell formats %
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' adjusting column sizes automatically
        ws.Columns("O:Q").AutoFit
     
    Next ws
End Sub
